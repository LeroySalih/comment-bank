# Permission and Role Requirements

## Overview

This document defines the role-based access control (RBAC) requirements for the
Comment Bank application.

## Roles

The application supports three distinct roles:

1. **Admin** - System administrators
2. **HOD** (Head of Department) - Department managers
3. **Teacher** - Teaching staff

### Multi-Role Support

Users may have multiple roles assigned. When a user has multiple roles, they
should have access to all routes permitted by any of their roles.

## Route Access Matrix

| Route Pattern   | Admin | HOD | Teacher | Unauthenticated         |
| --------------- | ----- | --- | ------- | ----------------------- |
| `/login`        | ✅    | ✅  | ✅      | ✅                      |
| `/` (Dashboard) | ✅    | ✅  | ✅      | ❌ Redirect to `/login` |
| `/admin/*`      | ✅    | ❌  | ❌      | ❌ Redirect to `/login` |
| `/admin/ccg`    | ✅    | ❌  | ❌      | ❌ Redirect to `/login` |
| `/hod/*`        | ❌    | ✅  | ❌      | ❌ Redirect to `/login` |
| `/class/*`      | ❌    | ❌  | ✅      | ❌ Redirect to `/login` |
| `/student/*`    | ❌    | ❌  | ✅      | ❌ Redirect to `/login` |

## Detailed Route Permissions

### Admin Routes (`/admin/*`)

- **Access**: Admin role ONLY
- **Purpose**: System administration, user management, subject creation,
  Common Comment Group (CCG) management
- **Includes**: `/admin/ccg` — Create, edit, delete common comment groups
  and their options; configure paragraph 2 wrapper template
- **Redirect**: Users without admin role → `/login`

### HOD Routes (`/hod/*`)

- **Access**: HOD role ONLY
- **Purpose**: Department management, subject management, comment bank creation
- **Redirect**: Users without HOD role → `/login`

### Teacher Routes (`/class/*`, `/student/*`)

- **Access**: Teacher role ONLY
- **Purpose**: Class management, student comment generation
- **Redirect**: Users without teacher role → `/login`

### Dashboard (`/`)

- **Access**: Any authenticated user
- **Purpose**: Landing page after login, shows role-appropriate content
- **Redirect**: Unauthenticated users → `/login`

## Navigation Menu Visibility

The navigation menu should display links based on the user's role(s):

### Admin Users See:

- Dashboard
- Admin Dashboard link (`/admin`)
- Common Comment Groups link (`/admin/ccg`)

### HOD Users See:

- Dashboard
- HOD Dashboard link (`/hod`)

### Teacher Users See:

- Dashboard
- Classes (links to their assigned classes)
- Students (links to students in their classes)

### Multi-Role Users See:

- Dashboard
- All links for each of their roles (combined)

## Authentication Requirements

### Login Page (`/login`)

- Accessible to all users (authenticated and unauthenticated)
- Accepts username and password
- On successful authentication, redirects to dashboard (`/`)

### Sign Out

- Available to all authenticated users
- Redirects to `/login` after sign out

### Session Management

- Sessions managed via NextAuth
- Token includes user roles for authorization checks
- Middleware validates token and roles on each protected route request

## Unauthorized Access Behavior

When a user attempts to access a route they don't have permission for:

1. **Redirect to `/login`**
2. No error message displayed (security best practice)
3. After login, user is NOT automatically redirected to the originally requested
   URL

## Test Users

For testing and development purposes:

| Username     | Password   | Role(s) |
| ------------ | ---------- | ------- |
| `admin`      | `password` | admin   |
| `leroysalih` | `password` | hod     |
| `teacher`    | `password` | teacher |

## Implementation Notes

### Middleware (`middleware.ts`)

- Uses NextAuth's `withAuth` middleware
- Checks route patterns and validates user roles
- Enforces strict role separation (no role has access to another role's routes)
- Redirects unauthorized users to `/login`

### Server Components

- Should perform additional role checks for sensitive operations
- Use `getServerSession()` to verify user authentication and roles
- Implement defense in depth (don't rely solely on middleware)

### Client Components

- Should hide UI elements based on user roles
- Use session data to conditionally render navigation links
- Remember: Client-side checks are for UX only, not security

## Security Considerations

1. **Defense in Depth**: Both middleware and server components should validate
   permissions
2. **Least Privilege**: Users only have access to routes necessary for their
   role
3. **No Information Disclosure**: Unauthorized access attempts should not reveal
   whether a route exists
4. **Session Security**: Tokens should be validated on every request
5. **Multi-Role Handling**: Users with multiple roles get combined permissions
   (union, not intersection)
