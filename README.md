# ExcelNinja 🥷

<img src="src/main/resources/logo.jpeg" alt="ExcelNinja Logo" width="200" />

**ExcelNinja** is a lightweight, modern Java library that makes Excel file handling as swift and precise as a ninja's blade. 
<br>Built on top of Apache POI, it provides a clean, annotation-driven API for seamless conversion between Java Object and Excel documents.

---

## Key Features

- **Simple & Intuitive**: One-line reads and writes with automatic type conversion
- **Annotation-Driven**: Map Excel columns to DTO fields using `@ExcelReadColumn` and `@ExcelWriteColumn`
- **Bidirectional**: Read Excel → DTO lists, Write DTO lists → Excel
- ️**Type-Safe**: Built-in validation with domain-driven value objects
- ️**Modern Architecture**: Clean hexagonal architecture with proper separation of concerns
- **Performance Logging**: Built-in performance metrics and detailed logging
- **Java 8+ Compatible**: Works with Java 8 and above

---

## Quick Start

### Gradle Dependency

```groovy
dependencies {
    implementation 'io.github.excel-ninja:excelNinja:0.0.7'
}
```

> **Zero Configuration Required**  
> ExcelNinja bundles all dependencies internally - no need to add Apache POI manually.

### Basic Usage

#### 1. Define Your DTO

```java
public class User {
    @ExcelReadColumn(headerName = "ID")
    @ExcelWriteColumn(headerName = "ID", order = 0)
    private Long id;

    @ExcelReadColumn(headerName = "Name")
    @ExcelWriteColumn(headerName = "Name", order = 1)
    private String name;

    @ExcelReadColumn(headerName = "Age")
    @ExcelWriteColumn(headerName = "Age", order = 2)
    private Integer age;

    @ExcelReadColumn(headerName = "Email")
    @ExcelWriteColumn(headerName = "Email", order = 3)
    private String email;

    // Constructors, getters, and setters...
}
```

#### 2. Read Excel to DTO List

```java
// From file path
List<User> users = NinjaExcel.read("users.xlsx", User.class);

// From File object
File file = new File("users.xlsx");
List<User> users = NinjaExcel.read(file, User.class);
```

#### 3. Write DTO List to Excel

```java
List<User> users = Arrays.asList(
    new User(1L, "Alice", 28, "alice@example.com"),
    new User(2L, "Bob", 32, "bob@example.com")
);

// Create document and write to file
ExcelDocument document = ExcelDocument.writer()
    .objects(users)
    .sheetName("Users")
    .create();

NinjaExcel.write(document, "output.xlsx");

// Or write to OutputStream
try (FileOutputStream out = new FileOutputStream("output.xlsx")) {
    NinjaExcel.write(document, out);
}
```

---

## Advanced Features

### Custom Type Conversion

ExcelNinja includes built-in support for common types including `LocalDate`:

```java
public class Employee {
    @ExcelReadColumn(headerName = "Hire Date")
    private LocalDate hireDate;
    
    @ExcelReadColumn(headerName = "Salary", defaultValue = "0")
    private Double salary;
    
    // Supports various date formats automatically:
    // yyyy-MM-dd, dd/MM/yyyy, MM/dd/yyyy, etc.
}
```

### Document Builder Pattern

For more control over Excel generation:

```java
ExcelDocument document = ExcelDocument.writer()
    .objects(employees)
    .sheetName("Employee Report")
    .columnWidth(0, 5000)  // Set column width
    .rowHeight(0, (short) 25)  // Set row height
    .create();
```

### Advanced Reading with Document API

```java
ExcelDocument document = NinjaExcel.read("data.xlsx", MyDto.class);

// Access document metadata
String sheetName = document.getSheetName().getValue();
int rowCount = document.getRowCount();
int columnCount = document.getColumnCount();

// Query specific cells
Object cellValue = document.getCellValue(0, "Name");
List<Object> ageColumn = document.getColumn("Age");

// Convert to entities
List<MyDto> entities = document.convertToEntities(MyDto.class, new DefaultConverter());
```

---

##  Architecture

ExcelNinja follows clean architecture principles:

```
📁 Application Layer
├── facade/
│   └── NinjaExcel.java          # Main API facade
└── port/
    └── ConverterPort.java       # Type conversion interface

📁 Domain Layer
├── model/                       # Value objects & domain models
│   ├── ExcelDocument.java      # Core document aggregate
│   ├── Headers.java            # Header management
│   ├── DocumentRows.java       # Row collection
│   └── SheetName.java          # Sheet name validation
├── annotation/                  # Column mapping annotations
└── exception/                   # Domain-specific exceptions

📁 Infrastructure Layer
├── io/                         # Apache POI adapters
├── converter/                  # Type conversion implementation
└── util/                       # Reflection utilities
```

---

## Annotation Reference

### `@ExcelReadColumn`

Maps DTO fields to Excel columns when reading:

```java
@ExcelReadColumn(
    headerName = "Full Name",     // Required: Excel column header
    type = String.class,          // Optional: Override field type
    defaultValue = "Unknown"      // Optional: Default if cell is empty
)
private String name;
```

### `@ExcelWriteColumn`

Maps DTO fields to Excel columns when writing:

```java
@ExcelWriteColumn(
    headerName = "Employee ID",   // Required: Excel column header
    order = 0                     // Optional: Column order (default: field order)
)
private Long id;
```

---

## Core Components

| Component | Description |
|-----------|-------------|
| **`NinjaExcel`** | Main facade providing `read()` and `write()` methods |
| **`ExcelDocument`** | Immutable document representation with builder patterns |
| **`@ExcelReadColumn`** | Annotation for Excel → DTO mapping |
| **`@ExcelWriteColumn`** | Annotation for DTO → Excel mapping |
| **`DefaultConverter`** | Handles type conversion between Excel and Java types |

---

## Performance & Logging

ExcelNinja provides detailed performance metrics:

```
[NINJA-EXCEL] Reading Excel file: users.xlsx (245.2 KB)
[NINJA-EXCEL] Successfully read 1000 records from users.xlsx in 150 ms (6666.67 records/sec)

[NINJA-EXCEL] Writing Excel document with 500 records to output.xlsx
[NINJA-EXCEL] Successfully wrote 500 records to output.xlsx (89.4 KB) in 95 ms (5263.16 records/sec)
```

---

##  Error Handling

ExcelNinja provides comprehensive error handling with descriptive messages:

```java
try {
    List<User> users = NinjaExcel.read("missing.xlsx", User.class);
} catch (DocumentConversionException e) {
    // Detailed error messages for:
    // - File not found
    // - Invalid Excel format
    // - Type conversion failures
    // - Missing headers
    // - Annotation configuration errors
}
```

---

## Testing

ExcelNinja includes comprehensive test coverage:

- ✅ Unit tests for all core components
- ✅ Integration tests for end-to-end workflows
- ✅ Value object validation tests
- ✅ Exception handling tests
- ✅ Performance benchmarks

Run tests with: `./gradlew test`

---

##  Contributing

We welcome contributions! ExcelNinja is designed to be:

- **Extensible**: Easy to add new type converters
- **Maintainable**: Clean architecture with proper separation
- **Testable**: Comprehensive test coverage
- **Documentation-first**: Clear APIs and examples

### Development Setup

```bash
git clone https://github.com/excel-ninja/excel-ninja-toolkit.git
cd excel-ninja-toolkit
./gradlew build
```

### Contributing Guidelines

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Add tests for your changes
4. Update documentation if needed
5. Ensure all tests pass (`./gradlew test`)
6. Submit a pull request

---

## 📝 License

This project is licensed under the Apache License 2.0 - see the [LICENSE](https://www.apache.org/licenses/LICENSE-2.0.txt) file for details.

---

## 🚀 Roadmap

- [ ] **Excel Templates**: Support for predefined Excel templates
- [ ] **Cell Styling**: Rich formatting and styling options
- [ ] **Multiple Sheets**: Read/write multiple sheets in one operation
- [ ] **Streaming**: Support for large files with streaming API
- [ ] **Custom Validators**: Field-level validation annotations
- [ ] **Excel Functions**: Support for Excel formulas and functions

---

##  Support

If you find ExcelNinja helpful, please consider:

-  Starring this repository
-  Reporting issues
-  Suggesting features
-  Contributing code

---

**Made by the ExcelNinja team**

*"Making Excel handling in Java as swift and precise as a ninja's blade"*