import "next-auth";
import "next-auth/jwt";

// NextAuth의 Session / JWT 타입에 우리가 추가로 저장하는 필드를 등록한다.
declare module "next-auth" {
  interface Session {
    accessToken?: string;
    error?: string;
  }
}

declare module "next-auth/jwt" {
  interface JWT {
    accessToken?: string;
    refreshToken?: string;
    expiresAt?: number;
    error?: string;
  }
}
