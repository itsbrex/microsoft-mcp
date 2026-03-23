# Device Code Flow Implementation Guide for Next.js with Better Auth

This guide documents how to implement Microsoft Device Code Flow authentication in a Next.js TypeScript application that uses Better Auth. It's based on the successful implementation in the `microsoft-mcp` project and the `inbox-zero` reference implementation.

## Overview

Device Code Flow (also called "device authorization flow") allows users to authenticate on devices that cannot easily display a browser, or in CLI/headless environments. The user is shown a code and URL, which they enter in a browser on any device to complete authentication.

### Key Characteristics

| Feature | Device Code Flow | Standard OAuth |
|---------|------------------|----------------|
| Browser required on device | No | Yes |
| User interaction | Code + URL | Redirect |
| Client secret needed | No (public client) | Yes |
| Best for | CLI, headless, TV apps | Web apps |
| Refresh token handling | Via MSAL cache | Via OAuth |

## Architecture

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────────┐
│   Client App    │────▶│   Next.js API    │────▶│  Microsoft Identity │
│ (shows code)    │     │   /device-code   │     │     Platform        │
└─────────────────┘     └──────────────────┘     └─────────────────────┘
         │                       │                         │
         │ 1. Initiate          │ 2. Get device code     │
         │◀──────────────────────│◀────────────────────────│
         │                       │                         │
         │ 3. User enters code at microsoft.com/devicelogin
         │                       │                         │
         │ 4. Poll for result   │ 5. Poll auth server    │
         │──────────────────────▶│─────────────────────────▶│
         │                       │                         │
         │ 6. Return tokens     │ 7. Return tokens        │
         │◀──────────────────────│◀────────────────────────│
```

## Prerequisites

1. **Node.js** >= 20.0.0
2. **Next.js** with App Router
3. **Better Auth** configured
4. **Prisma** (or similar ORM)
5. **@azure/msal-node** package

## Environment Variables

Add to your `.env` file:

```bash
# MSAL Device Code Flow Configuration
MSAL_ENABLED=true
MSAL_CLIENT_ID=d3590ed6-52b3-4102-aeff-aad2292ab01c  # Microsoft Office client ID (works out of box)
MSAL_TENANT_ID=common  # or your specific tenant ID
MSAL_DEBUG=false  # Set to true for verbose MSAL logging

# Required for token encryption
EMAIL_ENCRYPT_SECRET=your-encryption-secret-32-chars
EMAIL_ENCRYPT_SALT=your-encryption-salt

# Session management
AUTH_SECRET=your-auth-secret-for-better-auth
```

### Client ID Options

| Client ID | Use Case | Registration Required |
|-----------|----------|----------------------|
| `d3590ed6-52b3-4102-aeff-aad2292ab01c` | Microsoft Office (default) | No |
| Your Azure AD App ID | Custom permissions | Yes |

**Important**: The Microsoft Office client ID works out of the box for device code flow with pre-authorized Graph API permissions. No Azure app registration needed.

## Database Schema

Add the following field to your Account model (Prisma example):

```prisma
model Account {
  id                 String    @id @default(cuid())
  userId             String
  provider           String
  providerAccountId  String
  access_token       String?   @db.Text
  refresh_token      String?   @db.Text
  expires_at         DateTime?
  token_type         String?
  scope              String?

  // MSAL Device Code specific
  msal_cache         String?   @db.Text  // Encrypted MSAL token cache
  msal_cache_updated DateTime?

  @@unique([provider, providerAccountId])
}
```

Run migration:

```bash
pnpm prisma migrate dev --name add-msal-cache
```

## Implementation

### 1. Environment Configuration (`env.ts`)

```typescript
import { createEnv } from "@t3-oss/env-nextjs";
import { z } from "zod";

export const env = createEnv({
  server: {
    // Existing config...

    // MSAL Device Code Flow
    MSAL_CLIENT_ID: z.string().optional(),
    MSAL_TENANT_ID: z.string().optional(),
    MSAL_ENABLED: z.string().optional(), // "true" to enable
    MSAL_DEBUG: z.string().optional(), // "true" for verbose logging

    EMAIL_ENCRYPT_SECRET: z.string(),
    EMAIL_ENCRYPT_SALT: z.string(),
    AUTH_SECRET: z.string(),
  },
});
```

### 2. MSAL Device Code Module (`utils/outlook/msal-device-code.ts`)

```typescript
import {
  PublicClientApplication,
  type DeviceCodeRequest,
  type AuthenticationResult,
  type AccountInfo,
  type Configuration,
  LogLevel,
} from "@azure/msal-node";
import { env } from "@/env";

// Default Microsoft Office client ID - works without app registration
const MICROSOFT_OFFICE_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c";

// Use .default scope for device code flow
const MSAL_DEVICE_CODE_SCOPES = ["https://graph.microsoft.com/.default"];

// Feature flag check
export function isMSALDeviceCodeEnabled(): boolean {
  return env.MSAL_ENABLED === "true";
}

// MSAL configuration
export function getMSALConfig(): { clientId: string; tenantId: string } {
  return {
    clientId: env.MSAL_CLIENT_ID || MICROSOFT_OFFICE_CLIENT_ID,
    tenantId: env.MSAL_TENANT_ID || "common",
  };
}

// Cached MSAL app instance
let msalApp: PublicClientApplication | null = null;

export function getMSALApp(): PublicClientApplication {
  if (msalApp) return msalApp;

  const { clientId, tenantId } = getMSALConfig();

  const config: Configuration = {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    system: {
      loggerOptions: {
        loggerCallback: (level, message, containsPii) => {
          if (containsPii) return;
          if (env.MSAL_DEBUG === "true" || level === LogLevel.Error) {
            console.log(`MSAL [${LogLevel[level]}]: ${message}`);
          }
        },
        piiLoggingEnabled: false,
        logLevel: env.MSAL_DEBUG === "true" ? LogLevel.Verbose : LogLevel.Error,
      },
    },
  };

  msalApp = new PublicClientApplication(config);
  return msalApp;
}

// Active device code flows storage
interface ActiveFlow {
  deviceCode: string;
  userCode: string;
  verificationUri: string;
  expiresAt: Date;
  message: string;
  promise: Promise<AuthenticationResult | null>;
  resolve: (result: AuthenticationResult | null) => void;
  reject: (error: Error) => void;
}

const activeFlows = new Map<string, ActiveFlow>();

export interface DeviceCodeInitResponse {
  sessionId: string;
  userCode: string;
  verificationUri: string;
  expiresAt: Date;
  message: string;
}

export async function initiateDeviceCodeFlow(
  sessionId: string,
): Promise<DeviceCodeInitResponse> {
  if (!isMSALDeviceCodeEnabled()) {
    throw new Error("MSAL device code flow is not enabled");
  }

  const app = getMSALApp();

  let flowResolve: (result: AuthenticationResult | null) => void;
  let flowReject: (error: Error) => void;
  let capturedDeviceCode: { userCode: string; verificationUri: string; expiresIn: number; message: string } | null = null;

  const flowPromise = new Promise<AuthenticationResult | null>((resolve, reject) => {
    flowResolve = resolve;
    flowReject = reject;
  });

  const deviceCodeRequest: DeviceCodeRequest = {
    scopes: MSAL_DEVICE_CODE_SCOPES,
    deviceCodeCallback: (response) => {
      capturedDeviceCode = response;

      const expiresAt = new Date(Date.now() + response.expiresIn * 1000);

      activeFlows.set(sessionId, {
        deviceCode: response.deviceCode,
        userCode: response.userCode,
        verificationUri: response.verificationUri,
        expiresAt,
        message: response.message,
        promise: flowPromise,
        resolve: flowResolve!,
        reject: flowReject!,
      });
    },
    timeout: 900, // 15 minutes
  };

  // Start the flow (don't await - it blocks until user completes)
  const authPromise = app.acquireTokenByDeviceCode(deviceCodeRequest);

  // Wait for callback
  await new Promise<void>((resolve, reject) => {
    const checkInterval = setInterval(() => {
      if (capturedDeviceCode) {
        clearInterval(checkInterval);
        resolve();
      }
    }, 100);

    setTimeout(() => {
      clearInterval(checkInterval);
      if (!capturedDeviceCode) reject(new Error("Device code callback timeout"));
    }, 10_000);
  });

  const flow = activeFlows.get(sessionId);
  if (!flow || !capturedDeviceCode) {
    throw new Error("Failed to initiate device code flow");
  }

  // Wire up completion
  authPromise
    .then((result) => flow.resolve(result))
    .catch((error) => flow.reject(error));

  const expiresAt = new Date(Date.now() + capturedDeviceCode.expiresIn * 1000);

  return {
    sessionId,
    userCode: capturedDeviceCode.userCode,
    verificationUri: capturedDeviceCode.verificationUri,
    expiresAt,
    message: capturedDeviceCode.message,
  };
}

export interface PollResult {
  status: "pending" | "complete" | "expired" | "error";
  error?: string;
  result?: {
    accessToken: string;
    expiresAt: Date;
    scopes: string[];
    account: AccountInfo | null;
  };
}

export async function pollDeviceCodeFlow(sessionId: string): Promise<PollResult> {
  const flow = activeFlows.get(sessionId);

  if (!flow) {
    return { status: "expired" };
  }

  if (new Date() > flow.expiresAt) {
    activeFlows.delete(sessionId);
    return { status: "expired" };
  }

  // Non-blocking check
  const raceResult = await Promise.race([
    flow.promise
      .then((result) => ({ done: true as const, result, error: null }))
      .catch((error) => ({ done: true as const, result: null, error })),
    new Promise<{ done: false }>((resolve) =>
      setTimeout(() => resolve({ done: false }), 100),
    ),
  ]);

  if (!raceResult.done) {
    return { status: "pending" };
  }

  activeFlows.delete(sessionId);

  if (raceResult.error) {
    const errorMessage = raceResult.error instanceof Error
      ? raceResult.error.message
      : String(raceResult.error);

    if (errorMessage.includes("authorization_pending")) {
      return { status: "pending" };
    }

    return { status: "error", error: errorMessage };
  }

  if (!raceResult.result) {
    return { status: "error", error: "No authentication result" };
  }

  return {
    status: "complete",
    result: {
      accessToken: raceResult.result.accessToken,
      expiresAt: raceResult.result.expiresOn || new Date(Date.now() + 3_600_000),
      scopes: raceResult.result.scopes,
      account: raceResult.result.account,
    },
  };
}

export function cancelDeviceCodeFlow(sessionId: string): boolean {
  const flow = activeFlows.get(sessionId);
  if (flow) {
    flow.reject(new Error("Flow cancelled by user"));
    activeFlows.delete(sessionId);
    return true;
  }
  return false;
}
```

### 3. API Routes

#### Initiate Route (`app/api/outlook/device-code/initiate/route.ts`)

```typescript
import { NextResponse } from "next/server";
import { nanoid } from "nanoid";
import { initiateDeviceCodeFlow, isMSALDeviceCodeEnabled } from "@/utils/outlook/msal-device-code";

export async function POST() {
  if (!isMSALDeviceCodeEnabled()) {
    return NextResponse.json(
      { error: "MSAL device code authentication is not enabled" },
      { status: 403 },
    );
  }

  try {
    const sessionId = nanoid();
    const result = await initiateDeviceCodeFlow(sessionId);

    return NextResponse.json({
      sessionId: result.sessionId,
      userCode: result.userCode,
      verificationUri: result.verificationUri,
      expiresAt: result.expiresAt.toISOString(),
      message: result.message,
    });
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : "Unknown error";
    return NextResponse.json(
      { error: `Failed to initiate device code flow: ${errorMessage}` },
      { status: 500 },
    );
  }
}
```

#### Poll Route (`app/api/outlook/device-code/poll/route.ts`)

```typescript
import { type NextRequest, NextResponse } from "next/server";
import { z } from "zod";
import { serializeSignedCookie } from "better-call";
import { pollDeviceCodeFlow, isMSALDeviceCodeEnabled, getMSALApp } from "@/utils/outlook/msal-device-code";
import { encryptToken } from "@/utils/encryption";
import prisma from "@/utils/prisma";
import { env } from "@/env";

const pollRequestSchema = z.object({
  sessionId: z.string().min(1, "Session ID is required"),
});

export async function POST(request: NextRequest) {
  if (!isMSALDeviceCodeEnabled()) {
    return NextResponse.json(
      { error: "MSAL device code authentication is not enabled" },
      { status: 403 },
    );
  }

  try {
    const body = await request.json();
    const { sessionId } = pollRequestSchema.parse(body);

    const pollResult = await pollDeviceCodeFlow(sessionId);

    if (pollResult.status !== "complete" || !pollResult.result) {
      return NextResponse.json({ status: pollResult.status, error: pollResult.error });
    }

    // Authentication complete - create/update user and account
    const { accessToken, expiresAt, scopes, account } = pollResult.result;

    if (!account) {
      throw new Error("No account info in token response");
    }

    // Get user profile from Microsoft Graph
    const profileResponse = await fetch("https://graph.microsoft.com/v1.0/me", {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const profile = await profileResponse.json();

    const email = profile.mail?.toLowerCase() || profile.userPrincipalName?.toLowerCase();
    const name = profile.displayName;
    const providerAccountId = account.localAccountId;

    if (!email) {
      throw new Error("Could not determine user email");
    }

    // Encrypt access token
    const encryptedAccessToken = encryptToken(accessToken);

    // Find or create user
    let user = await prisma.user.findUnique({ where: { email } });

    if (!user) {
      user = await prisma.user.create({
        data: {
          email,
          name: name || email,
          emailVerified: true,
        },
      });
    }

    // Create or update account
    const accountRecord = await prisma.account.upsert({
      where: {
        provider_providerAccountId: {
          provider: "microsoft",
          providerAccountId,
        },
      },
      create: {
        userId: user.id,
        provider: "microsoft",
        providerAccountId,
        access_token: encryptedAccessToken,
        refresh_token: null, // MSAL handles refresh internally
        expires_at: expiresAt,
        token_type: "Bearer",
        scope: scopes.join(" "),
      },
      update: {
        access_token: encryptedAccessToken,
        expires_at: expiresAt,
        scope: scopes.join(" "),
      },
    });

    // Persist MSAL cache for future token refresh
    const app = getMSALApp();
    const tokenCache = app.getTokenCache();
    const serializedCache = tokenCache.serialize();

    if (serializedCache) {
      const encryptedCache = encryptToken(serializedCache);
      if (encryptedCache) {
        await prisma.account.updateMany({
          where: { provider: "microsoft", providerAccountId },
          data: {
            msal_cache: encryptedCache,
            msal_cache_updated: new Date(),
          },
        });
      }
    }

    // Create Better Auth session
    const sessionToken = crypto.randomUUID();
    const sessionExpires = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000); // 30 days

    await prisma.session.create({
      data: {
        sessionToken,
        userId: user.id,
        expires: sessionExpires,
      },
    });

    // Sign session token
    const secret = env.AUTH_SECRET;
    const signedToken = await serializeSignedCookie("", sessionToken, secret);

    // Build response with session cookie
    const isProduction = env.NODE_ENV === "production";
    const cookieName = isProduction
      ? "__Secure-better-auth.session_token"
      : "better-auth.session_token";

    const response = NextResponse.json({
      status: "complete",
      email,
      redirectUrl: "/welcome",
      message: "Authentication successful!",
    });

    response.cookies.set(cookieName, signedToken.replace(/^=/, ""), {
      httpOnly: true,
      secure: isProduction,
      sameSite: "lax",
      path: "/",
      expires: sessionExpires,
    });

    return response;
  } catch (error) {
    if (error instanceof z.ZodError) {
      return NextResponse.json(
        { status: "error", error: "Invalid request body" },
        { status: 400 },
      );
    }

    return NextResponse.json(
      { status: "error", error: "Poll failed" },
      { status: 500 },
    );
  }
}
```

### 4. MSAL Cache Persistence Plugin (`utils/outlook/msal-cache-plugin.ts`)

For production, persist the MSAL cache to survive server restarts:

```typescript
import type { ICachePlugin, TokenCacheContext } from "@azure/msal-node";
import prisma from "@/utils/prisma";
import { encryptToken, decryptToken } from "@/utils/encryption";

export function createPrismaCachePlugin(providerAccountId: string): ICachePlugin {
  return {
    async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
      const account = await prisma.account.findFirst({
        where: { provider: "microsoft", providerAccountId },
        select: { msal_cache: true },
      });

      if (account?.msal_cache) {
        const decrypted = decryptToken(account.msal_cache);
        if (decrypted) {
          cacheContext.tokenCache.deserialize(decrypted);
        }
      }
    },

    async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
      if (!cacheContext.cacheHasChanged) return;

      const serialized = cacheContext.tokenCache.serialize();
      const encrypted = encryptToken(serialized);

      if (encrypted) {
        await prisma.account.updateMany({
          where: { provider: "microsoft", providerAccountId },
          data: {
            msal_cache: encrypted,
            msal_cache_updated: new Date(),
          },
        });
      }
    },
  };
}
```

### 5. Token Refresh (`utils/outlook/client.ts`)

Integrate with your Outlook client for automatic token refresh:

```typescript
import { Client } from "@microsoft/microsoft-graph-client";
import { acquireMSALTokenSilent } from "./msal-device-code";
import { encryptToken } from "@/utils/encryption";
import prisma from "@/utils/prisma";

export async function getOutlookClientWithRefresh({
  accessToken,
  refreshToken,
  expiresAt,
  providerAccountId,
}: {
  accessToken?: string | null;
  refreshToken: string | null;
  expiresAt: number | null;
  providerAccountId?: string | null;
}) {
  const TOKEN_REFRESH_BUFFER_MS = 10 * 60 * 1000; // 10 minutes

  // Check if token is still valid
  if (accessToken && expiresAt && expiresAt > Date.now() + TOKEN_REFRESH_BUFFER_MS) {
    return createOutlookClient(accessToken);
  }

  // For device-code accounts (no refresh token), use MSAL silent acquisition
  if (!refreshToken && providerAccountId) {
    const msalResult = await acquireMSALTokenSilent(providerAccountId);

    if (!msalResult) {
      throw new Error("Authentication expired. Please log in again.");
    }

    // Update the account with fresh token
    await prisma.account.updateMany({
      where: { providerAccountId },
      data: {
        access_token: encryptToken(msalResult.accessToken),
        expires_at: msalResult.expiresAt,
      },
    });

    return createOutlookClient(msalResult.accessToken);
  }

  // Standard OAuth refresh flow...
  // (existing implementation)
}

function createOutlookClient(accessToken: string) {
  return Client.init({
    authProvider: (done) => done(null, accessToken),
  });
}
```

### 6. Frontend Component

```typescript
'use client';

import { useState, useEffect, useCallback } from "react";

interface DeviceCodeState {
  status: "idle" | "loading" | "awaiting_auth" | "complete" | "error";
  userCode?: string;
  verificationUri?: string;
  message?: string;
  error?: string;
  sessionId?: string;
}

export function DeviceCodeAuth() {
  const [state, setState] = useState<DeviceCodeState>({ status: "idle" });

  const initiateFlow = useCallback(async () => {
    setState({ status: "loading" });

    try {
      const response = await fetch("/api/outlook/device-code/initiate", {
        method: "POST",
      });

      if (!response.ok) throw new Error("Failed to initiate flow");

      const data = await response.json();

      setState({
        status: "awaiting_auth",
        userCode: data.userCode,
        verificationUri: data.verificationUri,
        message: data.message,
        sessionId: data.sessionId,
      });
    } catch (error) {
      setState({
        status: "error",
        error: error instanceof Error ? error.message : "Unknown error",
      });
    }
  }, []);

  // Poll for completion
  useEffect(() => {
    if (state.status !== "awaiting_auth" || !state.sessionId) return;

    const pollInterval = setInterval(async () => {
      try {
        const response = await fetch("/api/outlook/device-code/poll", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ sessionId: state.sessionId }),
        });

        const data = await response.json();

        if (data.status === "complete") {
          clearInterval(pollInterval);
          setState({ status: "complete", message: data.message });
          window.location.href = data.redirectUrl;
        } else if (data.status === "error" || data.status === "expired") {
          clearInterval(pollInterval);
          setState({ status: "error", error: data.error || "Flow expired" });
        }
      } catch {
        // Continue polling on network errors
      }
    }, 3000); // Poll every 3 seconds

    return () => clearInterval(pollInterval);
  }, [state.status, state.sessionId]);

  if (state.status === "idle") {
    return (
      <button onClick={initiateFlow} className="btn btn-primary">
        Sign in with Microsoft (Device Code)
      </button>
    );
  }

  if (state.status === "loading") {
    return <div>Initiating authentication...</div>;
  }

  if (state.status === "awaiting_auth") {
    return (
      <div className="p-4 border rounded">
        <h3>Sign in to Microsoft</h3>
        <p>{state.message}</p>
        <div className="my-4">
          <p>Go to: <a href={state.verificationUri} target="_blank" rel="noopener noreferrer" className="text-blue-500">{state.verificationUri}</a></p>
          <p className="text-2xl font-mono mt-2">Code: <strong>{state.userCode}</strong></p>
        </div>
        <p className="text-sm text-gray-500">Waiting for you to complete sign-in...</p>
      </div>
    );
  }

  if (state.status === "complete") {
    return <div>Authentication successful! Redirecting...</div>;
  }

  if (state.status === "error") {
    return (
      <div className="p-4 border border-red-300 rounded">
        <p className="text-red-600">Error: {state.error}</p>
        <button onClick={initiateFlow} className="mt-2 btn">Try Again</button>
      </div>
    );
  }

  return null;
}
```

## Key Implementation Details

### 1. Token Scope: `.default`

Device code flow uses `https://graph.microsoft.com/.default` scope:
- Requests all permissions the app has been pre-authorized for
- No user consent screen (permissions already granted to Microsoft Office client)
- Works with the default Microsoft Office client ID

### 2. No Refresh Token Exposure

MSAL device code flow doesn't directly expose the refresh token:
- Refresh is handled internally by MSAL via the token cache
- The cache must be persisted for token refresh after server restarts
- Use the `ICachePlugin` interface for database persistence

### 3. Account Identifier

The user's identity comes from the token response, not from input:
- `account.localAccountId` - unique identifier for the Microsoft account
- User email is fetched from Microsoft Graph `/me` endpoint
- The `MICROSOFT_MCP_ACCOUNT_ID` env var is only for local file naming in CLI scenarios

### 4. Better Auth Integration

Device code creates sessions compatible with Better Auth:
- Creates user and account records in the database
- Creates a signed session cookie that Better Auth recognizes
- Uses the same session model and cookie configuration

## Comparison: microsoft-mcp vs inbox-zero Implementation

| Aspect | microsoft-mcp | inbox-zero |
|--------|---------------|------------|
| Storage | File-based tokens | Database (Prisma) |
| Cache persistence | Not needed (CLI) | Encrypted in DB |
| Session handling | N/A (MCP server) | Better Auth sessions |
| Token refresh | Manual via HTTP | MSAL silent + fallback |
| Multi-account | File per account | DB records per user |

## Troubleshooting

### "MSAL device code flow is not enabled"
- Set `MSAL_ENABLED=true` in environment

### "Failed to acquire token silently"
- MSAL cache may be corrupted or expired
- User needs to re-authenticate
- Check `msal_cache` field is being persisted

### "fetch failed" errors
- Node.js 22+ has undici connection pool issues
- Implement custom `INetworkModule` with `keepalive: false`
- Add timeout handling with AbortController

### Token not refreshing
- Ensure `msal_cache` is encrypted and stored after auth
- Check `msal_cache_updated` timestamp
- Verify cache plugin is correctly loading/saving

## Security Considerations

1. **Token Encryption**: Always encrypt tokens before storage
2. **Cache Encryption**: MSAL cache contains sensitive credentials
3. **HTTPS**: Use secure cookies in production
4. **Validation**: Validate all user input and poll requests
5. **Session Security**: Use signed cookies with proper flags

## References

- [Microsoft Device Code Flow Documentation](https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code)
- [MSAL Node.js Device Code Sample](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/samples/msal-node-samples/device-code)
- [Better Auth Documentation](https://better-auth.com)
- [microsoft-mcp Source Code](https://github.com/itsbrex/microsoft-mcp)
- [inbox-zero Source Code](https://github.com/elie222/inbox-zero)
