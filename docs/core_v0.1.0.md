# AnyFile Core v0.1.0 – Project Overview

## Project Name
- `AnyFile`

## Version
- `0.1.0` (Core Foundation preview)

## Scope
- Establish the universal foundation for file handling operations that downstream modules—Excel, PDF, Word, CSV, and more—extend.
- Define the base API shape, type definitions, and extensibility points needed by all future AnyFile packages.

## Goal
- Deliver a lightweight, type-safe, and consistent API that developers can rely on when working with any kind of file through modular packages under the `@anyfile/*` namespace.
- Provide a cohesive developer experience that hides format-specific complexity behind a unified contract.


## Core Components
- `AnyFile` static facade exposing `register`, `open`, and handler introspection hooks.
- Handler `registry` with collision protection for file types and extensions, plus detection hooks for non-extension sources.
- Shared `fileTypes` definitions (`FileType`, `FileMetadata`) powering type-safety across the ecosystem.
- Utility helpers under `core/src/utils` for path and metadata normalization.


## Base API (v0.1.0)
- `AnyFile.register(handler)` – registers pluggable handlers contributed by submodules.
- `AnyFile.open(source, options?)` – infers handler from extension or detection hook, returning a rich file instance (`read`, `write`, `convert`).
- `AnyFile.getHandler(type)` – retrieves registered handlers for diagnostic or composition flows.
- Registry management helpers (`clearRegistry`, `listRegisteredHandlers`) exposed for internal tooling and tests.


## Tooling
- pnpm workspace scaffold with root TypeScript configuration and Vitest test harness.
- `tsup` build target from the core package for dual ESM/CJS output and type declarations.

