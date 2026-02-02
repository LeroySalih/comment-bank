# Security Review - Comment Bank Application

**Review Date:** 2026-02-01\
**Reviewer:** AI Security Analysis\
**Application:** Comment Bank (Next.js 16.1.2, PostgreSQL, NextAuth.js)

---

## Executive Summary

This security review identifies **10 critical and high-priority security
issues** that should be addressed before deploying to production. The
application has a solid foundation with proper authentication, authorization,
and data encryption, but requires improvements in several areas including
environment variable management, security headers, rate limiting, and CSRF
protection.

**Overall Security Rating:** ⚠️ **MODERATE** (Requires improvements before
production deployment)

---

## 🔴 Critical Issues

### 1. Weak NEXTAUTH_SECRET

**Severity:** 🔴 CRITICAL\
**Location:** `.env` line 2

**Issue:**

```env
NEXTAUTH_SECRET="supersecret123"
```

The `NEXTAUTH_SECRET` is only 14 characters and uses a predictable value. This
secret is used to sign and encrypt JWT tokens and session data.

**Risk:**

- Session hijacking
- JWT forgery
- Unauthorized access to user accounts

**Recommendation:**

```bash
# Generate a strong secret (minimum 32 characters)
openssl rand -base64 32
```

Update `.env`:

```env
NEXTAUTH_SECRET="<generated-strong-secret-here>"
```

---

### 2. Duplicate DATABASE_URL Configuration

**Severity:** 🔴 CRITICAL\
**Location:** `.env` lines 1 and 8

**Issue:**

```env
DATABASE_URL="file:./dev.db"  # Line 1 - SQLite
DATABASE_URL="postgresql://postgres:your-super-secret-and-long-postgres-password@localhost:5432/comment_bank"  # Line 8 - PostgreSQL
```

Two conflicting `DATABASE_URL` entries exist. The second one will override the
first, but this creates confusion and potential security risks.

**Risk:**

- Accidental connection to wrong database
- Exposure of database credentials in version control
- Configuration drift between environments

**Recommendation:**

1. Remove the SQLite entry (line 1)
2. Ensure `.env` is in `.gitignore`
3. Use `.env.example` for documentation:

```env
# .env.example
DATABASE_URL="postgresql://username:password@host:port/database"
NEXTAUTH_SECRET="<generate-with-openssl-rand-base64-32>"
NEXTAUTH_URL="http://localhost:3001"
PUPIL_ENCRYPTION_KEY="<64-character-hex-string>"
```

---

### 3. Database Credentials in Plain Text

**Severity:** 🔴 CRITICAL\
**Location:** `.env` line 8

**Issue:** The PostgreSQL password is stored in plain text in the `.env` file:

```env
DATABASE_URL="postgresql://postgres:your-super-secret-and-long-postgres-password@localhost:5432/comment_bank"
```

**Risk:**

- If `.env` is accidentally committed to version control, credentials are
  exposed
- Local file access = database access
- No rotation mechanism

**Recommendation:**

1. **Verify `.gitignore` includes `.env`:**

```gitignore
.env
.env.local
.env.*.local
```

2. **For production, use environment-specific secrets management:**
   - **Vercel:** Use Environment Variables in dashboard
   - **AWS:** Use AWS Secrets Manager or Parameter Store
   - **Docker:** Use Docker secrets
   - **Kubernetes:** Use Kubernetes secrets

3. **Implement credential rotation policy**

---

## 🟠 High Priority Issues

### 4. Missing Security Headers

**Severity:** 🟠 HIGH\
**Location:** `next.config.ts`

**Issue:** The Next.js configuration lacks security headers to protect against
common web vulnerabilities.

**Risk:**

- Clickjacking attacks (no X-Frame-Options)
- MIME-type sniffing (no X-Content-Type-Options)
- XSS attacks (no Content-Security-Policy)
- Information disclosure (no X-Powered-By removal)

**Recommendation:**

Update `next.config.ts`:

```typescript
import type { NextConfig } from "next";

const nextConfig: NextConfig = {
    headers: async () => [
        {
            source: "/:path*",
            headers: [
                {
                    key: "X-DNS-Prefetch-Control",
                    value: "on",
                },
                {
                    key: "Strict-Transport-Security",
                    value: "max-age=63072000; includeSubDomains; preload",
                },
                {
                    key: "X-Frame-Options",
                    value: "SAMEORIGIN",
                },
                {
                    key: "X-Content-Type-Options",
                    value: "nosniff",
                },
                {
                    key: "X-XSS-Protection",
                    value: "1; mode=block",
                },
                {
                    key: "Referrer-Policy",
                    value: "strict-origin-when-cross-origin",
                },
                {
                    key: "Permissions-Policy",
                    value: "camera=(), microphone=(), geolocation=()",
                },
                {
                    key: "Content-Security-Policy",
                    value:
                        "default-src 'self'; script-src 'self' 'unsafe-eval' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; img-src 'self' data: https:; font-src 'self' data:;",
                },
            ],
        },
    ],
    poweredByHeader: false,
};

export default nextConfig;
```

---

### 5. No Rate Limiting

**Severity:** 🟠 HIGH\
**Location:** Authentication endpoints, server actions

**Issue:** There is no rate limiting on:

- Login attempts (`/api/auth/[...nextauth]`)
- Server actions (admin, HOD actions)
- User creation endpoints

**Risk:**

- Brute force attacks on login
- Denial of Service (DoS)
- Resource exhaustion
- Credential stuffing attacks

**Recommendation:**

**Option 1: Use Upstash Rate Limit (Recommended for Vercel)**

```bash
npm install @upstash/ratelimit @upstash/redis
```

Create `lib/rate-limit.ts`:

```typescript
import { Ratelimit } from "@upstash/ratelimit";
import { Redis } from "@upstash/redis";

// Create a rate limiter that allows 5 requests per 10 seconds
export const loginRateLimit = new Ratelimit({
    redis: Redis.fromEnv(),
    limiter: Ratelimit.slidingWindow(5, "10 s"),
    analytics: true,
});

// Create a rate limiter for server actions (10 requests per minute)
export const actionRateLimit = new Ratelimit({
    redis: Redis.fromEnv(),
    limiter: Ratelimit.slidingWindow(10, "1 m"),
    analytics: true,
});
```

**Option 2: Use next-rate-limit (Simpler, in-memory)**

```bash
npm install next-rate-limit
```

---

### 6. Missing CSRF Protection for State-Changing Operations

**Severity:** 🟠 HIGH\
**Location:** Server actions

**Issue:** While Next.js Server Actions have some built-in CSRF protection,
there's no explicit CSRF token validation for critical state-changing
operations.

**Risk:**

- Cross-Site Request Forgery attacks
- Unauthorized actions performed on behalf of authenticated users

**Recommendation:**

NextAuth.js provides CSRF protection for authentication flows, but for
additional security on server actions:

1. **Leverage Next.js built-in protection** (already in place via Server
   Actions)
2. **Add explicit checks for critical operations:**

```typescript
// lib/csrf.ts
import { headers } from "next/headers";

export async function validateOrigin() {
    const headersList = headers();
    const origin = headersList.get("origin");
    const host = headersList.get("host");

    if (origin && !origin.includes(host || "")) {
        throw new Error("Invalid origin");
    }
}

// Use in critical server actions:
export const deleteSubject = withRole("admin", async (subjectId: string) => {
    await validateOrigin(); // Add this
    // ... rest of the code
});
```

---

### 7. Insufficient Password Policy

**Severity:** 🟠 HIGH\
**Location:** `lib/validation-schemas.ts` line 11

**Issue:**

```typescript
password: z.string().min(6, 'Password must be at least 6 characters'),
```

Minimum password length of 6 characters is too weak for modern security
standards.

**Risk:**

- Weak passwords susceptible to brute force
- Dictionary attacks
- Credential stuffing

**Recommendation:**

Update `lib/validation-schemas.ts`:

```typescript
password: z.string()
  .min(12, 'Password must be at least 12 characters')
  .regex(/[A-Z]/, 'Password must contain at least one uppercase letter')
  .regex(/[a-z]/, 'Password must contain at least one lowercase letter')
  .regex(/[0-9]/, 'Password must contain at least one number')
  .regex(/[^A-Za-z0-9]/, 'Password must contain at least one special character'),
```

**Additional Recommendations:**

- Implement password strength meter on UI
- Check against common password lists (e.g., Have I Been Pwned API)
- Enforce password rotation policy for admin accounts

---

## 🟡 Medium Priority Issues

### 8. Legacy Unencrypted Data Support

**Severity:** 🟡 MEDIUM\
**Location:** `lib/encryption.ts` lines 47-50

**Issue:**

```typescript
if (!isEncrypted(encryptedText)) {
    // For now, return as-is to support migration
    // TODO: Remove this once all data is encrypted
    return encryptedText;
}
```

The decrypt function returns unencrypted data as-is for backward compatibility.

**Risk:**

- Unencrypted PII may remain in database
- Inconsistent data protection
- Compliance violations (GDPR, FERPA)

**Recommendation:**

1. **Create a migration script:**

```typescript
// scripts/encrypt-legacy-data.ts
import { prisma } from "@/lib/prisma";
import { encrypt, isEncrypted } from "@/lib/encryption";

async function encryptLegacyData() {
    const pupils = await prisma.pupil.findMany();

    for (const pupil of pupils) {
        const updates: any = {};

        if (!isEncrypted(pupil.firstName)) {
            updates.firstName = encrypt(pupil.firstName);
        }
        if (!isEncrypted(pupil.lastName)) {
            updates.lastName = encrypt(pupil.lastName);
        }

        if (Object.keys(updates).length > 0) {
            await prisma.pupil.update({
                where: { admissionNumber: pupil.admissionNumber },
                data: updates,
            });
        }
    }

    console.log("Migration complete");
}

encryptLegacyData();
```

2. **After migration, remove fallback:**

```typescript
export function decrypt(encryptedText: string): string {
    // Remove the isEncrypted check and always expect encrypted data
    if (!isEncrypted(encryptedText)) {
        throw new Error("Data is not encrypted - migration required");
    }
    // ... rest of decryption logic
}
```

---

### 9. No Session Timeout Configuration

**Severity:** 🟡 MEDIUM\
**Location:** `app/api/auth/[...nextauth]/route.ts`

**Issue:** No explicit session timeout or maxAge configuration for JWT tokens.

**Risk:**

- Sessions may persist indefinitely
- Increased window for session hijacking
- Compliance issues (some regulations require session timeouts)

**Recommendation:**

Update `app/api/auth/[...nextauth]/route.ts`:

```typescript
export const authOptions: NextAuthOptions = {
    session: {
        strategy: "jwt",
        maxAge: 8 * 60 * 60, // 8 hours
        updateAge: 60 * 60, // Update session every hour
    },
    jwt: {
        maxAge: 8 * 60 * 60, // 8 hours
    },
    // ... rest of config
};
```

---

### 10. Missing Audit Logging

**Severity:** 🟡 MEDIUM\
**Location:** Throughout application

**Issue:** While there is logging via the `Logger` utility, there's no
comprehensive audit trail for security-sensitive operations:

- User login/logout
- Role changes
- Data access (especially PII)
- Failed authentication attempts
- Administrative actions

**Risk:**

- Inability to detect security breaches
- No forensic trail for investigations
- Compliance violations (GDPR requires audit logs)

**Recommendation:**

1. **Create audit logging system:**

```typescript
// lib/audit-log.ts
import { prisma } from "./prisma";
import { logger } from "./logger";

export enum AuditAction {
    USER_LOGIN = "USER_LOGIN",
    USER_LOGOUT = "USER_LOGOUT",
    USER_LOGIN_FAILED = "USER_LOGIN_FAILED",
    USER_CREATED = "USER_CREATED",
    USER_ROLE_CHANGED = "USER_ROLE_CHANGED",
    PUPIL_VIEWED = "PUPIL_VIEWED",
    PUPIL_UPDATED = "PUPIL_UPDATED",
    PUPIL_DELETED = "PUPIL_DELETED",
    SUBJECT_CREATED = "SUBJECT_CREATED",
    SUBJECT_DELETED = "SUBJECT_DELETED",
}

interface AuditLogEntry {
    action: AuditAction;
    userId?: string;
    username?: string;
    ipAddress?: string;
    userAgent?: string;
    resourceType?: string;
    resourceId?: string;
    details?: Record<string, any>;
    success: boolean;
}

export async function createAuditLog(entry: AuditLogEntry) {
    try {
        // Log to console/file immediately
        logger.info("Audit Log", entry);

        // Store in database for long-term retention
        // You'll need to create an AuditLog model in Prisma schema
        // await prisma.auditLog.create({ data: entry })
    } catch (error) {
        // Never let audit logging break the application
        logger.error("Failed to create audit log", { error, entry });
    }
}
```

2. **Add to Prisma schema:**

```prisma
model AuditLog {
  id           String   @id @default(cuid())
  action       String
  userId       String?
  username     String?
  ipAddress    String?
  userAgent    String?
  resourceType String?
  resourceId   String?
  details      Json?
  success      Boolean
  createdAt    DateTime @default(now())

  @@index([userId])
  @@index([action])
  @@index([createdAt])
}
```

3. **Integrate into server actions:**

```typescript
export const updateUserRoles = withRole("admin", async (
    userId: string,
    roleNames: string[],
) => {
    try {
        await userRepository.updateRoles(userId, roleNames);

        await createAuditLog({
            action: AuditAction.USER_ROLE_CHANGED,
            userId: session.user.id,
            username: session.user.username,
            resourceType: "User",
            resourceId: userId,
            details: { newRoles: roleNames },
            success: true,
        });

        return { success: true };
    } catch (error) {
        await createAuditLog({
            action: AuditAction.USER_ROLE_CHANGED,
            userId: session.user.id,
            username: session.user.username,
            resourceType: "User",
            resourceId: userId,
            details: { error: error.message },
            success: false,
        });

        return handleServerActionError(error);
    }
});
```

---

## ✅ Security Strengths

The application demonstrates several security best practices:

### 1. **Strong Password Hashing**

- ✅ Uses bcrypt with cost factor of 12
- ✅ Proper async hashing implementation
- ✅ Secure password comparison

### 2. **Robust Authorization**

- ✅ Role-based access control (RBAC)
- ✅ Middleware protection on routes
- ✅ `withRole` HOC for server actions
- ✅ Granular permission checks (admin, HOD, teacher)

### 3. **Data Encryption**

- ✅ AES-256-GCM for PII encryption
- ✅ Proper IV generation (random per encryption)
- ✅ Authentication tags for integrity
- ✅ Automatic encryption/decryption in repositories

### 4. **Input Validation**

- ✅ Comprehensive Zod schemas
- ✅ Server-side validation for all inputs
- ✅ Type-safe validation with TypeScript

### 5. **SQL Injection Prevention**

- ✅ Prisma ORM with parameterized queries
- ✅ No raw SQL queries found
- ✅ Type-safe database operations

### 6. **XSS Prevention**

- ✅ React's automatic escaping
- ✅ No `dangerouslySetInnerHTML` in application code
- ✅ No `eval()` usage

### 7. **Error Handling**

- ✅ Standardized error classes
- ✅ Centralized error handler
- ✅ No sensitive data in error messages
- ✅ Proper logging without exposing internals

---

## 📋 Security Checklist for Production

Before deploying to production, ensure:

- [ ] **Environment Variables**
  - [ ] Generate strong `NEXTAUTH_SECRET` (32+ characters)
  - [ ] Remove duplicate `DATABASE_URL` entries
  - [ ] Verify `.env` is in `.gitignore`
  - [ ] Create `.env.example` without secrets
  - [ ] Use secrets management in production (not `.env` files)

- [ ] **Security Headers**
  - [ ] Implement all recommended headers in `next.config.ts`
  - [ ] Test with [securityheaders.com](https://securityheaders.com)

- [ ] **Rate Limiting**
  - [ ] Implement rate limiting on login endpoint
  - [ ] Add rate limiting to server actions
  - [ ] Configure appropriate limits based on usage patterns

- [ ] **Password Policy**
  - [ ] Update minimum password length to 12 characters
  - [ ] Add complexity requirements
  - [ ] Implement password strength indicator

- [ ] **Data Encryption**
  - [ ] Run migration script to encrypt legacy data
  - [ ] Remove fallback for unencrypted data
  - [ ] Verify all PII is encrypted at rest

- [ ] **Session Management**
  - [ ] Configure session timeout (8 hours recommended)
  - [ ] Implement session refresh mechanism
  - [ ] Add logout functionality

- [ ] **Audit Logging**
  - [ ] Create AuditLog model
  - [ ] Implement audit logging for sensitive operations
  - [ ] Set up log retention policy

- [ ] **HTTPS**
  - [ ] Enforce HTTPS in production
  - [ ] Configure HSTS header
  - [ ] Redirect HTTP to HTTPS

- [ ] **Database Security**
  - [ ] Use strong database passwords
  - [ ] Restrict database access by IP
  - [ ] Enable database encryption at rest
  - [ ] Regular database backups

- [ ] **Dependency Security**
  - [ ] Run `npm audit` and fix vulnerabilities
  - [ ] Keep dependencies up to date
  - [ ] Use `npm audit fix` regularly

---

## 🔍 Additional Recommendations

### 1. **Implement Content Security Policy (CSP)**

Fine-tune the CSP header to be more restrictive:

```typescript
'Content-Security-Policy': "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; font-src 'self'; connect-src 'self'; frame-ancestors 'none';"
```

### 2. **Add Security Testing**

- Integrate OWASP ZAP or similar security scanner
- Add security-focused E2E tests
- Regular penetration testing

### 3. **Implement Multi-Factor Authentication (MFA)**

For admin accounts, consider adding:

- TOTP (Time-based One-Time Password)
- SMS verification
- Email verification

### 4. **Database Connection Pooling**

Configure Prisma connection pooling for production:

```typescript
// prisma/schema.prisma
datasource db {
  provider = "postgresql"
  url      = env("DATABASE_URL")
  directUrl = env("DIRECT_URL") // For migrations
}
```

### 5. **Monitoring and Alerting**

Set up monitoring for:

- Failed login attempts
- Unusual data access patterns
- Server errors
- Performance metrics

---

## 📚 References

- [OWASP Top 10](https://owasp.org/www-project-top-ten/)
- [Next.js Security Best Practices](https://nextjs.org/docs/app/building-your-application/configuring/security-headers)
- [NextAuth.js Security](https://next-auth.js.org/configuration/options#security)
- [NIST Password Guidelines](https://pages.nist.gov/800-63-3/sp800-63b.html)
- [GDPR Compliance](https://gdpr.eu/)

---

## Conclusion

The Comment Bank application has a solid security foundation with proper
authentication, authorization, and data encryption. However, **critical issues
with environment variable management and missing security headers must be
addressed before production deployment**.

**Priority Actions:**

1. 🔴 Generate strong `NEXTAUTH_SECRET`
2. 🔴 Clean up `.env` file and verify `.gitignore`
3. 🟠 Implement security headers
4. 🟠 Add rate limiting
5. 🟠 Strengthen password policy

**Timeline Recommendation:**

- **Critical issues:** Fix immediately (before any production deployment)
- **High priority:** Fix within 1 week
- **Medium priority:** Fix within 1 month

---

**Review Status:** ✅ Complete\
**Next Review:** Recommended after implementing fixes and before production
deployment
