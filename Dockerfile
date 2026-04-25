FROM node:20-alpine AS base

# ── deps: install production + dev dependencies ──────────────────────────────
FROM base AS deps
RUN apk add --no-cache libc6-compat
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci

# ── builder: compile the Next.js app ─────────────────────────────────────────
FROM base AS builder
WORKDIR /app
COPY --from=deps /app/node_modules ./node_modules
COPY . .

# Dummy values so `next build` doesn't fail env validation at build time.
# Real values are injected at runtime via docker-compose / your platform.
ENV DATABASE_URL=postgresql://placeholder:placeholder@localhost:5432/placeholder
ENV NEXTAUTH_SECRET=placeholder_secret_for_build_only_32chars
ENV NEXTAUTH_URL=http://localhost:3003
ENV PUPIL_ENCRYPTION_KEY=0000000000000000000000000000000000000000000000000000000000000000
ENV NODE_ENV=production

RUN npm run build

# ── runner: lean production image ────────────────────────────────────────────
FROM base AS runner
WORKDIR /app
ENV NODE_ENV=production

RUN addgroup --system --gid 1001 nodejs && \
    adduser  --system --uid 1001 nextjs

COPY --from=builder /app/public ./public
COPY --from=builder /app/.next  ./.next
COPY --from=builder /app/node_modules ./node_modules
COPY --from=builder /app/package.json ./package.json

USER nextjs
EXPOSE 3003

CMD ["node_modules/.bin/next", "start", "-p", "3003"]
