# Comment Bank

A Next.js application for teachers to generate and manage student report comments.

## Getting Started

First, run the development server:

```bash
npm run dev
# or
yarn dev
# or
pnpm dev
# or
bun dev
```

Open [http://localhost:3000](http://localhost:3000) with your browser to see the result.

## Claude Code Skills

### `/report-comment` — Report Comment Generator

Generates a set of three levelled school report comment templates (High, Medium, Low) for a given subject and topic.

**Usage:**
```
/report-comment Computing, Theoretic Knowledge
```

Or invoke without arguments and Claude will prompt for Subject and Topic.

**How it works:**
1. Claude generates a draft High/Medium/Low comment set using standard template variables
2. You review and request adjustments per level if needed
3. Once approved, the final comments are displayed as a table and saved to:

```
comments/<Subject>/<Topic>.md
```

**Template variables available in comments:**

| Variable | Meaning |
|----------|---------|
| `<Name>` | Pupil's name |
| `<he/she>` | Gender pronoun (subject) |
| `<his/her>` | Gender pronoun (possessive) |
| `<him/her>` | Gender pronoun (object) |
| `<Subject>` | Subject name |
| `<Year>` | Academic year |
| `<EoYLevel>` | End of year level achieved |
| `<TargetLevel>` | Target level |

## Learn More

- [Next.js Documentation](https://nextjs.org/docs)
- [Learn Next.js](https://nextjs.org/learn)
