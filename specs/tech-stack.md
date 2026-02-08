# Comment Bank - Technology Stack & Architecture

## Overview

Comment Bank is a web application designed for teachers to efficiently create
and manage student report comments. The system supports role-based access
control with three user types: Admin, Head of Department (HOD), and Teacher.

---

## Technology Stack

### Frontend

#### Core Framework

- **Next.js 16.1.2** - React framework with App Router
  - Server-side rendering (SSR)
  - Server components for data fetching
  - Client components for interactivity
  - File-based routing in `/app` directory

#### UI & Styling

- **React 19.2.3** - UI library
- **Tailwind CSS 4** - Utility-first CSS framework
  - Custom design tokens in `globals.css`
  - Dark mode support
- **Google Fonts (Inter)** - Typography
- **Material Symbols** - Icon library
- **Lucide React** - Additional icon set

#### UI Components & Libraries

- **@hello-pangea/dnd** - Drag-and-drop functionality
- **clsx** & **tailwind-merge** - Conditional class name utilities

### Backend

#### Runtime & Framework

- **Node.js** - JavaScript runtime
- **Next.js API Routes** - Backend API endpoints
  - Located in `/app/api`
  - Server actions in `/app/actions.ts`

#### Authentication

- **NextAuth.js 4.24.13** - Authentication solution
  - Credentials provider for username/password login
  - JWT-based sessions
  - Role-based middleware protection
  - Custom callbacks for role management

#### Database

- **PostgreSQL** - Primary database
- **Prisma 6.19.2** - ORM and database toolkit
  - Type-safe database client
  - Schema-first development
  - Migration management
  - Seeding support

#### Security

- **bcryptjs** - Password hashing
- **Middleware-based route protection** - Role-based access control
- **Zod 4.3.6** - Runtime type validation and schema validation

### Development & Testing

#### Testing

- **Playwright 1.57.0** - End-to-end testing
  - Chrome-only testing for speed and reliability
  - Permission, authentication, and admin flow tests
- **Vitest** - Unit testing framework (configured but not actively used)

#### Development Tools

- **TypeScript 5** - Type safety
- **ESLint 9** - Code linting
- **tsx** - TypeScript execution for scripts

#### Build Tools

- **PostCSS** - CSS processing
- **Tailwind CSS PostCSS plugin** - CSS optimization

---

## Architecture

### Application Structure

```
comment-bank-claude/
├── app/                    # Next.js App Router
│   ├── admin/             # Admin dashboard & management
│   │   └── ccg/           # Common Comment Group management
│   ├── hod/               # HOD subject & comment management
│   ├── class/             # Teacher class management
│   ├── student/           # Student assignment & comments
│   ├── login/             # Authentication
│   ├── api/               # API routes
│   │   └── auth/          # NextAuth configuration
│   ├── actions.ts         # Server actions
│   ├── layout.tsx         # Root layout
│   └── page.tsx           # Dashboard
├── components/            # Reusable React components
├── lib/                   # Utility functions & helpers
│   ├── access-control.ts  # Role checking utilities
│   └── prisma.ts          # Prisma client singleton
├── prisma/                # Database schema & migrations
│   ├── schema.prisma      # Database schema
│   └── seed.ts            # Database seeding
├── specs/                 # Documentation
│   └── permissions.md     # Permission requirements
├── tests/                 # E2E tests
│   └── permissions.spec.ts
└── middleware.ts          # NextAuth middleware
```

### Data Model

#### Core Entities

**User Management:**

- `User` - Application users with roles
- `Role` - Admin, HOD, Teacher roles

**Academic Structure:**

- `Subject` - Academic subjects (e.g., Computer Science)
- `Class` - Class groups within subjects
- `Pupil` - Students

**Comment System:**

- `CommentGroup` - Subject-specific groups of comment options (e.g., "Written Production")
- `CommentOption` - Predefined comment templates with codes (H/M/L)
- `CommonCommentGroup` - Common comment groups applied to all subjects (e.g., "Academic Performance", "Effort")
- `CommonCommentOption` - Predefined comment templates for common groups
- `CommentParagraphTemplate` - Admin-configurable wrapper template for combining CCGs into paragraph 2
- `Assignment` - Links pupils to classes; includes a `linkedData` (JSONB) field
  for flexible data-driven comment groups
- `PupilCode` - Selected comment codes per pupil per group (subject-specific)
- `CommonPupilCode` - Selected comment codes per pupil per common group (per assignment)

**Linked Comment Groups:**

Both `CommentGroup` and `CommonCommentGroup` support an optional linked mode:
- `isLinked` (Boolean, default false) - Whether this group is auto-populated
- `linkedField` (String, nullable) - The key in `Assignment.linkedData` to match
  against (e.g., `"behaviour"`, `"effort"`, `"homework"`)

#### Relationships

```mermaid
erDiagram
    User ||--o{ Role : has
    User ||--o{ Subject : manages
    User ||--o{ Class : teaches
    Subject ||--o{ Class : contains
    Subject ||--o{ CommentGroup : has
    Class ||--o{ Assignment : has
    Pupil ||--o{ Assignment : enrolled
    Assignment ||--o{ PupilCode : has
    Assignment ||--o{ CommonPupilCode : has
    CommentGroup ||--o{ CommentOption : contains
    CommentGroup ||--o{ PupilCode : references
    CommonCommentGroup ||--o{ CommonCommentOption : contains
    CommonCommentGroup ||--o{ CommonPupilCode : references
```

### Authentication & Authorization

#### Authentication Flow

1. User submits credentials via `/login`
2. NextAuth validates against database
3. JWT token issued with user ID and roles
4. Token stored in session cookie

#### Authorization Layers

**1. Middleware (Edge Protection)**

- File: `middleware.ts`
- Runs before page loads
- Route-based role checking:
  - `/admin/*` → Admin only
  - `/hod/*` → HOD only
  - `/class/*`, `/student/*` → Teacher only
  - `/` → Any authenticated user
- Redirects unauthorized users to `/login`

**2. Server Components (Data Access)**

- Use `getServerSession()` to check roles
- Filter database queries based on user permissions
- Example: Teachers only see their assigned classes

**3. UI Components (UX)**

- Conditional rendering based on roles
- Hide navigation links for unauthorized routes
- Utility functions in `lib/access-control.ts`

### Key Features

#### Role-Based Dashboards

**Admin (`/admin`)**

- User management
- Role assignment
- System-wide oversight
- Common Comment Group (CCG) management (`/admin/ccg`)
  - Create, edit, delete common comment groups and their options
  - Configure paragraph 2 wrapper template

**HOD (`/hod`)**

- Subject management
- Comment group creation
- Comment template editing
- Department-level reporting

**Teacher (`/class`, `/student`)**

- Class roster viewing
- Student assignment management
- Comment selection and customization
- Report generation

#### Comment Generation System

There are two types of comment groups: **Common Comment Groups (CCGs)** that
apply to all subjects, and **Subject-Specific Comment Groups** that belong to
individual subjects. Either type can optionally be configured as a **linked
group**, where the selected option is auto-populated from uploaded student data.

##### Common Comment Groups (CCGs)

CCGs are managed by **admins only** and apply across all subjects. The initial
set is: Academic Performance, Behaviour, Homework, Effort, and Overall. Admins
can add, edit, and remove CCGs over time.

- CCG option texts are **shared globally** — one set of templates for all
  subjects. Variables like `<Subject>` handle per-subject differentiation.
- CCG templates support all existing variables (`<Name>`, `<He>`, `<Subject>`,
  `<TargetLevel>`, `<EoYLevel>`, etc.)
- Teachers select CCG codes **per student per subject** (i.e., a student can
  receive different effort ratings in different subjects).
- Admin UI located at `/admin/ccg`.

##### Subject-Specific Comment Groups

Managed by **HODs** as before:

1. **Template Creation (HOD)**
   - Create comment groups (e.g., "Written Production")
   - Define comment options with codes (H/M/L)
   - Use variables: `<Name>`, `<Subject>`, `<TargetLevel>`, `<EoYLevel>`

2. **Comment Selection (Teacher)**
   - Select code for each comment group per student
   - View real-time comment preview
   - Edit final comment text

##### Linked Comment Groups

Linked comment groups are a special type of comment group whose selected option
is **auto-populated from a data field** on the student's assignment record. They
can be either CCGs or subject-specific groups.

**Key Characteristics:**

- **Read-only for teachers** — the auto-selected option cannot be changed by the
  teacher. The group is displayed with a visual indicator showing it was
  auto-filled and is locked.
- Teachers can still edit the **final comment text** in the textarea. The
  existing lock/revert flow applies: editing the final comment locks all comment
  banks; reverting clears the edit and unlocks them.
- **Admin-managed** — admins create and configure linked groups for both CCGs and
  subject-specific groups.

**Data Storage:**

- A new `linkedData` column (JSONB) on the `Assignment` model stores flexible
  key-value data uploaded via the pupil data import.
- Initial fields: `behaviour`, `effort`, `homework` — each with possible values:
  `A*`, `A`, `B`, `C`, `D`, `E`, `F`, `NA`.
- The system is designed for **extensibility**: admins can define additional
  linkable fields in future without schema migrations. Available fields are
  inferred from the columns present in the uploaded pupil data.

**Admin Configuration:**

1. Create a comment group (name, title, paragraph position) as normal.
2. Mark the group as **"Linked"** and select which data field it links to (e.g.
   `behaviour`). The available fields are those present in the uploaded data.
3. Add options where the **code must exactly match** a possible value of the
   linked data field (e.g. codes: `A*`, `A`, `B`, `C`, `D`, `E`, `F`, `NA`).

**Auto-Selection Logic:**

- On page load, the system reads the student's `linkedData` value for the
  configured field and matches it to an option code.
- If a match is found, that option is auto-selected and displayed.
- If **no match is found** (unmatched value), a **per-student warning** is shown
  and that student's comment **cannot be edited or copied** to the clipboard.

**Example:**

The existing Effort CCG could become a linked group tied to the `effort` field:

| Code | Text |
|------|------|
| `A*` | `<He> consistently demonstrates exceptional effort and engagement in <Subject>.` |
| `A` | `<He> consistently puts in excellent effort in lessons.` |
| `B` | `<He> generally puts in good effort in lessons.` |
| ... | ... |
| `NA` | `Effort data is not available for <Name>.` |

If a student's uploaded data contains `effort: "A"`, the `A` option is
automatically selected and locked.

##### Final Comment Structure

The generated comment follows a **fixed 4-paragraph layout**:

| Paragraph | Source | Content |
|-----------|--------|---------|
| **P1** | CCG | Academic Performance selected text |
| **P2** | CCG (combined) | Effort + Behaviour + Homework, joined via an admin-configurable **wrapper template** |
| **P3** | Subject-specific | Existing subject comment groups (joined as before) |
| **P4** | CCG | Overall selected text |

**Paragraph 2 Wrapper Template:**
- Configured by admins at `/admin/ccg`.
- Uses placeholders such as `<Effort>`, `<Behaviour>`, `<Homework>` that are
  replaced with the selected option text for each CCG.
- Example: `"<Effort> <Behaviour> <Homework>"` — the system substitutes each
  placeholder with the teacher's selected option text, then applies standard
  variable replacement (`<Name>`, `<He>`, etc.) to the result.

##### Variable Replacement

- Gender-based pronoun replacement (`<He>`, `<His>`, `<him>`, etc.)
- Student name insertion (`<Name>`)
- Subject-specific text insertion (`<Subject>`)
- Level data interpolation (`<TargetLevel>`, `<EoYLevel>`)
- Applied to both CCG and subject-specific comment texts

##### Teacher UX (CommentEditor)

Comment groups are displayed to the teacher **interleaved by paragraph order**:

1. Academic Performance (CCG)
2. Effort, Behaviour, Homework (CCGs)
3. Subject-specific groups (in display order)
4. Overall (CCG)

### Deployment Configuration

#### Environment Variables

- `DATABASE_URL` - PostgreSQL connection string
- `NEXTAUTH_SECRET` - JWT signing secret
- `NEXTAUTH_URL` - Application base URL
- `SITE_NAME` - Application name for redirects

#### Scripts

- `npm run dev` - Development server (port 3001)
- `npm run build` - Production build with Prisma generation
- `npm start` - Production server
- `npx tsx prisma/seed.ts` - Seed database

---

## Security Considerations

### Implemented

✅ Password hashing with bcrypt\
✅ JWT-based session management\
✅ Role-based middleware protection\
✅ Server-side authorization checks\
✅ SQL injection prevention (Prisma ORM)\
✅ Type-safe database queries

### Best Practices

- Passwords never stored in plaintext
- Sensitive routes protected at multiple layers
- User input validated with Zod schemas
- CSRF protection via NextAuth
- Secure session cookies (httpOnly, secure flags)

---

## Performance Optimizations

- **Server Components** - Reduce client-side JavaScript
- **Prisma Connection Pooling** - Efficient database connections
- **Next.js Image Optimization** - Automatic image optimization
- **Static Generation** - Where applicable
- **Code Splitting** - Automatic via Next.js

---

## Testing Strategy

### E2E Tests (Playwright)

Tests run against **Chrome only** (`--project=chromium`) for speed and reliability.

#### Test Files

| File | Coverage |
|------|----------|
| `tests/auth.spec.ts` | Login/logout flows, role-based redirect |
| `tests/permissions.spec.ts` | 24 scenarios — Admin, HOD, Teacher, unauthenticated access |
| `tests/hod-management.spec.ts` | HOD subject and comment group management |
| `tests/teacher-comment-flow.spec.ts` | Teacher class/student comment selection |
| `tests/admin-dashboard.spec.ts` | Admin dashboard tabs (Users, Subjects, Classes, Deadlines, Activity Log) |
| `tests/admin-ccg.spec.ts` | CCG CRUD — create/edit/delete groups and options, wrapper template |

#### Shared Helpers

`tests/helpers.ts` exports:
- `login(page, username, password)` — reusable login flow
- `TEST_USERS` — credentials for admin, hod, teacher, teacher2, teacher3

#### Test Database

- **Database:** `comment_bank_test` (configured in `.env.test`)
- **Reset:** `npm run db:test:reset` — drops, migrates, and seeds via `prisma/seed.ts`
- **Safety:** The reset script aborts if `DATABASE_URL` doesn't contain `comment_bank_test`

#### Running Tests

```bash
# Reset test DB and run all E2E tests
npm run test:e2e:setup

# Run E2E tests (assumes test DB is already seeded)
npm run test:e2e

# Run E2E tests with Playwright UI
npm run test:e2e:ui

# Reset test DB only
npm run db:test:reset
```

### Test Configuration

- **Chrome only** — Firefox/WebKit removed for faster runs
- Sequential execution (`workers: 1`) to avoid session conflicts
- `webServer` block auto-starts dev server on port 3001 with `.env.test`
- Headless mode for CI/CD
- HTML reports for debugging
- Test users seeded in database via `prisma/seed.ts`

---

## Future Considerations

### Data Dump (CSV Export)

Admin users can export every database table as CSV files to avoid vendor lock-in.

- **Route**: `/admin/data-dump`
- **Access**: Admin only
- **Behaviour**: Single button triggers a server action that queries all tables,
  converts each to CSV, and bundles them into a ZIP download.
- **Tables exported**: User (without passwords), Role, Subject, Class, Pupil
  (decrypted names), Assignment, CommentGroup, CommentOption, PupilCode,
  CommonCommentGroup, CommonCommentOption, CommonPupilCode, AppSetting,
  Deadline, AuditLog.
- **File format**: UTF-8 CSV with header row, one file per table, delivered as
  `comment-bank-export-YYYY-MM-DD.zip`.
- **Security**: Passwords are excluded from User export. Pupil first/last names
  are decrypted before export so the CSV contains readable data.

### Potential Enhancements

- **Caching** - Redis for session storage
- **File Storage** - S3 for document uploads
- **Email** - Report distribution via email
- **Analytics** - Usage tracking and reporting
- **API** - RESTful API for integrations
- **Mobile** - Progressive Web App (PWA) support

### Scalability

- Horizontal scaling via load balancer
- Database read replicas
- CDN for static assets
- Containerization (Docker)
- Kubernetes orchestration
