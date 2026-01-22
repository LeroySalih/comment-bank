import { withAuth } from "next-auth/middleware"

export default withAuth({
  callbacks: {
    authorized: ({ req, token }) => {
      const path = req.nextUrl.pathname
      if (path.startsWith("/admin")) {
        return token?.roles?.includes("admin") ?? false
      }
      if (path.startsWith("/hod")) {
        return (token?.roles?.includes("hod") || token?.roles?.includes("admin")) ?? false
      }
      return !!token
    }
  },
  pages: {
    signIn: "/login",
  },
})

export const config = {
  matcher: ["/", "/class/:path*", "/student/:path*", "/admin/:path*", "/hod/:path*"]
}
