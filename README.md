# Access VBA Library

A modular collection of reusable classes for Access VBA development.

This repository provides common components such as data access helpers, utility functions, and debugging tools to simplify and standardize VBA projects.

---

## Features

* Reusable class-based architecture
* Transaction management and safe query execution
* Utility components for common tasks
* Interface-driven design for flexibility
* Debug and testing support

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
