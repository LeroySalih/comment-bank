import NextAuth from "next-auth"
import CredentialsProvider from "next-auth/providers/credentials"
import { pool } from "@/lib/db"
import { compare } from "bcryptjs"
import { logAuthEvent } from "@/lib/audit-log"

import { NextAuthOptions } from "next-auth"

export const authOptions: NextAuthOptions = {

  trustHost: true,
  session: {
    strategy: "jwt",
  },
  pages: {
    signIn: "/login",
  },
  providers: [
    CredentialsProvider({
      name: "Credentials",
      credentials: {
        username: { label: "Username", type: "text" },
        password: { label: "Password", type: "password" }
      },
      async authorize(credentials) {
        if (!credentials?.username || !credentials.password) {
          return null
        }

        const { rows: userRows } = await pool.query(
          `SELECT u.*, json_agg(json_build_object('id', r.id, 'name', r.name)) FILTER (WHERE r.id IS NOT NULL) as roles
           FROM "User" u
           LEFT JOIN "_RoleToUser" ru ON ru."B" = u.id
           LEFT JOIN "Role" r ON r.id = ru."A"
           WHERE u.username = $1
           GROUP BY u.id`,
          [credentials.username]
        )

        const userRow = userRows[0] ?? null
        const user = userRow ? {
          ...userRow,
          Role: userRow.roles ?? []
        } : null

        if (!user) {
          // Log failed sign-in attempt - user not found
          logAuthEvent('sign_in_failed', undefined, credentials.username, {
            reason: 'user_not_found'
          })
          return null
        }

        // Check if user is active
        if (!user.isActive) {
          logAuthEvent('sign_in_failed', user.id, credentials.username, {
            reason: 'user_inactive'
          })
          return null
        }

        const isPasswordValid = await compare(credentials.password, user.password)

        if (!isPasswordValid) {
          // Log failed sign-in attempt
          logAuthEvent('sign_in_failed', user.id, credentials.username, {
            reason: 'invalid_password'
          })
          return null
        }

        // Log successful sign-in
        logAuthEvent('sign_in', user.id, user.username, {
          roles: user.Role.map((r: any) => r.name)
        })

        return {
          id: user.id,
          name: user.username,
          username: user.username,
          roles: user.Role.map((r: any) => r.name)
        }
      }
    })
  ],
  events: {
    async signOut({ token }) {
      // Log sign-out
      if (token?.id) {
        logAuthEvent('sign_out', token.id as string, token.username as string)
      }
    }
  },
  callbacks: {
    async jwt({ token, user }) {
      if (user) {
        token.id = user.id
        token.username = (user as any).username
        token.roles = (user as any).roles
      }
      return token
    },
    async session({ session, token }) {
      if (session.user) {
        session.user.id = token.id
        session.user.username = token.username
        session.user.roles = token.roles as any
      }
      return session
    }
  }
}

const handler = NextAuth(authOptions)

export { handler as GET, handler as POST }
