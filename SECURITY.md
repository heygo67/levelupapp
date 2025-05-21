# Security Policy for Student Level-Up Checker

This application is designed to process Excel enrollment reports to identify students ready to level up. It is a stateless, file-processing tool hosted via Render. The app was built with privacy and security best practices in mind.

---

### Security Features

### 1. Stateless, In-Memory Processing

- No data is stored or persisted across requests.
- Uploaded files are processed in memory using `Roo`, and immediately discarded after the response is generated.
- No sessions or cookies are used.

### 2. Input Validation

- Only files with `.xlsx` or `.xls` extensions are allowed.
- Filenames are restricted to alphanumeric characters, underscores, hyphens, and periods.
- File size is limited to 5MB to prevent denial-of-service via large uploads.

### 3. Security Headers

The app sets the following HTTP security headers to reduce attack surface:

- `Content-Security-Policy`: Enforces `default-src 'self'` to prevent XSS.
- `Strict-Transport-Security`: Forces HTTPS on supporting browsers.
- `X-Content-Type-Options: nosniff`: Prevents MIME sniffing.
- `X-Frame-Options: DENY`: Blocks clickjacking attempts.

### 4. Rack::Protection Middleware

`Rack::Protection` is used to guard against:

- Cross-Site Scripting (XSS)
- HTTP header injection
- Path traversal
- Other common Rack-based web threats

### 5. File Integrity Logging

Each uploaded file is hashed with SHA-256 and logged (alongside client IP and timestamp) for audit purposes — without storing any user data:
