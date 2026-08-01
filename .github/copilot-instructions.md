# DevHelper - Copilot Instructions

## Project Overview

DevHelper is a VBA library that enables Git-friendly development for VBA projects.

The primary goal is to make VBA development feel similar to modern software development:

- Export VBA source files into a repository
- Manage them with Git
- Review and edit them using modern editors such as VS Code
- (Future) Import them back into VBA

This project values maintainability and simplicity over feature completeness.

### About `Dev*` modules

Modules prefixed with `Dev` are part of the DevHelper library itself.

They are maintained by the library developer and are not intended to be modified by library users.

Some `Public` procedures exist only to support DevHelper's own development workflow (for example, exporting or updating DevHelper modules).

The visibility (`Public` / `Private`) of a procedure does not necessarily indicate that it is part of the library's public API.

Some `Public` procedures exist only because they must be callable from the VBA editor during DevHelper maintenance.

When reviewing the project, distinguish between: 

- Public APIs intended for library users
- Public procedures intended only for DevHelper maintenance


---

## Design Principles

### Prefer small, incremental changes

Small, verifiable improvements are preferred over ambitious refactoring.

Think:

> "Wait. Don't rush. This is Zhuge Liang's trap."
> ___（まて、あわてるな。これは孔明の罠だ。）___

Implement one feature at a time.

---

### Respect single responsibility

Each procedure should have one clear responsibility.

Public procedures should provide simple APIs or entry points required for DevHelper maintenance.

Private procedures should implement detailed logic.

---

### Keep dependencies simple

Avoid circular dependencies.

The current dependency direction is: 

`DevBootstrap`

↑

`DevProjIO`

`DevBootstrap` provides shared constants, helper functions, and DevHelper maintenance utilities.

`DevProjIO` provides repository import/export functionality for library users.

Do not introduce dependencies in the opposite direction.

Although both modules expose `Public` procedures, they serve different audiences and should not be treated as a single public API surface.

Do not recommend changing a procedure's visibility solely to make the API appear cleaner. First consider its intended audience.

Review each module according to its intended role rather than applying generic library design principles.

---

### Prefer readability

Readable code is preferred over clever code.

Choose descriptive names.

Avoid unnecessary abbreviations.

Keep procedures reasonably short.

---

## Error Handling

Use `RaiseError()` for programmer errors.

Recoverable failures (such as individual export failures) should be reported without stopping the entire export process whenever practical.

---

## Review Priorities

When reviewing code, prioritize:

1. API design
2. Responsibility separation
3. Future maintainability
4. Consistency with existing code
5. VBA best practices

Do not recommend Python-, Java-, or C#-style solutions unless they fit naturally in VBA.

Prefer solutions that fit naturally within VBA and the VBIDE object model.

### Intent over convention

Some design choices in this project intentionally differ from common software design practices due to VBA/VBE limitations.

When suggesting improvements, first understand the reason behind the current design before recommending structural changes.

---

## Coding Style

### General

- Use explicit, descriptive names.
- Preserve the existing coding style unless there is a strong reason to change it.
- Write comments to explain **why**, not **what**, whenever possible.
- Add comments when intent is not obvious.

### Visibility

- Keep Public APIs minimal.
- Prefer Private helper procedures.
- Expose only procedures intended to be used by library users.

### Naming

- Prefix module-level variables with `m_`.
- Prefix procedure arguments with `a_`.
- Use `UPPER_SNAKE_CASE` for constants.

### Procedure Calls

- Use the `Call` keyword for procedure calls that do not use a return value.
- Always use parentheses when passing arguments to procedure calls.
- Preserve this style consistently throughout the project.

### Control Flow

- Prefer guard clauses and early exits over deep nesting.
- Since VBA has no `Continue` statement, use a `Continue:` label for loop continuation when appropriate.

### Dependencies

- Prefer late binding (`CreateObject`) to avoid unnecessary library references.
- When using library constants would require a reference, define equivalent constants locally instead.

### Refactoring
 
- Avoid premature abstraction.
- Refactor only when repetition becomes meaningful.
- Prefer small, incremental improvements over large rewrites.

### Code Organization

- Place Public APIs before Private helper procedures.
- Use bookmark procedures such as `Private Sub AA_HelperFunctions(): End Sub` to organize the Procedure List. These bookmark procedures are intentional and should not be removed.

### Developer experiments

Procedures under `AA_Experiments` are intentional manual verification helpers.

They are used to confirm behavior directly in the VBA environment and should not be removed as dead code without review.

Due to VBA's module-level access restrictions, some manual verification provedures are kept in the same module as the procedures they verify.

Do not move these procedures to separate modules unless the required access design is considered

### Consistency

When modifying existing code, prefer consistency with the surrounding code over introducing a different style or pattern.

---

## Long-term Vision

DevHelper is intended to become a practical toolkit for Git-based VBA development.

Design decisions should favor long-term maintainability rather than short-term convenience.

---

## Philosophy

This project intentionally avoids over-engineering.

Whenever multiple designs seem possible, prefer the simpler solution that solves today's problem well.

Future refactoring is acceptable when new requirements appear.