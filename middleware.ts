import { withAuth } from "next-auth/middleware"

export default withAuth({
  callbacks: {
    authorized: ({ req, token }) => {
      const path = req.nextUrl.pathname
      
      // Admin routes - admin role ONLY
      if (path.startsWith("/admin")) {
        return token?.roles?.includes("admin") ?? false
      }
      
      // HOD routes - HOD role ONLY
      if (path.startsWith("/hod")) {
        return token?.roles?.includes("hod") ?? false
      }
      
      // Teacher routes - teacher role ONLY
      if (path.startsWith("/class") || path.startsWith("/student")) {
        return token?.roles?.includes("teacher") ?? false
      }
      
      // Dashboard and other routes - any authenticated user
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
