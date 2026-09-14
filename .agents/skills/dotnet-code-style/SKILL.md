---
name: dotnet-code-style
description: Apply an explicit-flow coding style when writing or modifying C# or VB.NET code. Use the current project for architecture, naming, formatting, and business rules not covered by this skill.
---

# Explicit-flow .NET style

Make the execution path visible through straightforward control flow, meaningful local variables, and cohesive operations. Apply these rules to newly written or directly modified code without restructuring unrelated code.

## Defined outcomes

- Write application code with non-throwing control flow: validate inputs before use and represent every handled outcome through return values.
- Return a boolean or established sentinel when all non-success outcomes require the same behavior. Use the project's established status or result type when callers must distinguish outcomes.
- Use non-throwing APIs for operations that can fail during normal execution.

## Code shape

- Use `var` for initialized C# local variables and inferred `Dim` declarations in VB.
- Give meaningful transformation stages their own local variables.
- Keep simple expressions direct. Introduce a local when it names a real stage or makes complicated logic easier to understand.
- Prefer guard clauses and sequential flow: return early from invalid, unusable, or completed branches, then continue without an `else`.
- Return a simple final expression directly when an additional variable would add no meaning.
- Use block control flow for conditional choices. Use a complete `if`/`else` block only when both branches are necessary to express one mutually exclusive decision. Apply the same rule with block `If...Then...Else` in VB.
- Break long conditions into named boolean variables, then combine or return those variables.
- Keep intent in names and structure.
