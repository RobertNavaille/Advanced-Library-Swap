# AI Implementation Guide & System Context

This document outlines the architecture and requirements for the "Swap Library" Figma plugin, updated to support dynamic library connections via Share Links.

## Project Overview
A Figma plugin that allows users to swap components and styles between different libraries. The system is moving away from hardcoded library references ("Monkey"/"Shark") to a fully dynamic system where users can connect any Figma library.

## Core Requirements

### 1. Dynamic Library Connection
- **Input Method**: Users connect a library by pasting a **Figma Share Link** (e.g., `https://www.figma.com/design/fileKey/...`).
- **Parsing**: The plugin must extract the `fileKey` from the provided URL.
- **Metadata Fetching**:
  - Use the **Figma REST API** (`GET /v1/files/:fileKey`) to fetch library details (name, last modified, etc.).
  - Requires a **Personal Access Token (PAT)** provided by the user.
- **Storage**: Library metadata is stored locally using `figma.clientStorage`.

### 2. Data Persistence
- **Connected Libraries**: Stored in `clientStorage` under key `connected_libraries`.
- **Access Token**: Stored in `clientStorage` under key `figma_access_token`.
- **Structure**:
  ```typescript
  interface ConnectedLibrary {
    id: string;          // The fileKey
    name: string;        // Library Name
    type: 'Remote';
    lastSynced: string;  // ISO Date
    // Potentially cache component/style keys here to avoid constant re-fetching
    components?: Record<string, string>; // Name -> Key
    styles?: Record<string, string>;     // Name -> Key
  }
  ```

### 3. Refactored Swap Logic
- **Goal**: Remove all hardcoded "Monkey"/"Shark" logic.
- **Flow**:
  1. User selects a **Source Library** and a **Target Library** from their connected list.
  2. Plugin scans the selection for instances/styles from the Source Library.
  3. Plugin looks up the corresponding asset in the Target Library by **Name**.
  4. Plugin performs the swap using `swapComponent` or `swapStyle`.

### 4. User Interface
- **Add Library Modal**: Accepts a Share Link.
- **Settings**: Input field for Figma Personal Access Token.
- **Library List**: Displays connected libraries with options to remove or re-sync.

## Technical Stack
- **Language**: TypeScript (`code.ts`)
- **UI**: HTML/CSS/JS (`ui.html`)
- **API**: Figma Plugin API + Figma REST API (via `fetch`)

## Implementation Plan
1.  **Link Parsing**: Implement regex to extract `fileKey` from standard Figma URLs.
2.  **API Integration**: Implement `fetchLibraryMetadata(fileKey, token)` in `code.ts`.
3.  **Storage Logic**: Update `saveConnectedLibraries` and `loadConnectedLibraries` to handle the new data structure.
4.  **Swap Refactor**: Rewrite `performLibrarySwap` to iterate through the dynamic `CONNECTED_LIBRARIES` list instead of static maps.

## Current Status
- [x] UI updated to accept Share Links.
- [ ] `handleAddLibraryByLink` needs implementation in `code.ts`.
- [ ] Swap logic still contains hardcoded references.
