# Build stage
FROM golang:1.21-alpine AS builder

WORKDIR /app

# Install build dependencies
RUN apk add --no-cache git

# Copy go mod files
COPY go.mod go.sum ./
RUN go mod download

# Copy source code
COPY . .

# Build the application
RUN CGO_ENABLED=0 GOOS=linux go build -o main .

# Final stage with LibreOffice for PDF conversion
FROM alpine:latest

# Install LibreOffice and dependencies for PDF conversion
RUN apk add --no-cache \
    libreoffice \
    openjdk11-jre \
    ttf-dejavu \
    fontconfig \
    && fc-cache -f

WORKDIR /app

# Copy binary and template from builder
COPY --from=builder /app/main .
COPY template.xlsx .

# Expose port (Render will set PORT env var)
EXPOSE 8080

# Run the application
CMD ["./main"]
