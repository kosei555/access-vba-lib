# Access VBA Library

A modular collection of reusable classes for Microsoft Access VBA development.
Simplifies database operations, improves code maintainability, and standardizes development practices.

---

## Features

- Promote modular design in traditionally monolithic VBA codebases
- Standardize database access patterns
- Reduce boilerplate and repetitive DAO code
- Encourage safer data operations with transaction support
- Improve long-term maintainability of Access applications

---

## Getting Started

1. Clone or download this repository
2. Import required `.cls` / `.bas` files into your Access VBA project
3. Use only the components you need

---

## Example

```vba
Dim qm As New QueryManager

qm.RegisterQuery "Q_InsertUser"
qm.SetParam "Q_InsertUser", "name", "test"

qm.BeginTrans
qm.ExecQuery "Q_InsertUser"
qm.CommitTrans
```

---

## Project Structure

* `services` — Main entry classes
* `adapters` — External dependencies (DB, IO, etc.)
* `interfaces` — Abstractions
* `debug` — Debugging utilities

---

## Philosophy

This project aims to:

* Improve code reusability in Access VBA
* Encourage modular and maintainable design
* Provide practical building blocks for real-world usage

---

## Status

Work in progress. Contributions are welcome.

---

## License

MIT License
