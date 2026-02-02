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
  - Cross-browser testing (Chromium, Firefox, WebKit)
  - Permission and authentication tests
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

- `CommentGroup` - Groups of comment options (e.g., "Written Production")
- `CommentOption` - Predefined comment templates with codes (H/M/L)
- `Assignment` - Links pupils to classes
- `PupilCode` - Selected comment codes per pupil per group

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
    CommentGroup ||--o{ CommentOption : contains
    CommentGroup ||--o{ PupilCode : references
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

1. **Template Creation (HOD)**
   - Create comment groups (e.g., "Written Production")
   - Define comment options with codes (H/M/L)
   - Use variables: `<Name>`, `<Subject>`, `<TargetLevel>`, `<EoYLevel>`

2. **Comment Selection (Teacher)**
   - Select code for each comment group per student
   - View real-time comment preview
   - Edit final comment text

3. **Variable Replacement**
   - Gender-based name replacement
   - Subject-specific text insertion
   - Level data interpolation

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

- **Permission Tests** - 24 scenarios covering all roles
- **Authentication Tests** - Login/logout flows
- **Navigation Tests** - Menu and direct URL access
- **Cross-browser** - Chromium, Firefox, WebKit

### Test Configuration

- Sequential execution to avoid session conflicts
- Headless mode for CI/CD
- HTML reports for debugging
- Test users seeded in database

---

## Future Considerations

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
