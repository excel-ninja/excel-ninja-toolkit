# ExcelNinja

<img src="src/main/resources/logo.jpeg" alt="ExcelNinja Logo" width="200" />

**ExcelNinja** is a lightweight Java library that simplifies reading from and writing to Excel files.  
Built on top of Apache POI, it allows easy conversion between DTOs and Excel documents using annotations.

As an open-source project, ExcelNinja welcomes contributions and improvements.  
If you find any issues or have ideas to make it better, feel free to open a pull request or issue.  
Together, let's make Excel handling in Java sharper and faster — like a "ninja".

---

## Supported Java Versions

- **Java 8 and above**  
  (ExcelNinja is fully compatible with Java 8 syntax and runtime.)

---

## Features

- Read Excel files into DTO lists using `NinjaExcel.read`
- Write DTO lists into Excel files using `NinjaExcel.write`
- Column mapping via annotations (`@ExcelReadColumn`, `@ExcelWriteColumn`)
- Supports basic type conversion including `LocalDate`
- Customizable type conversion via `ConverterPort` interface

---

## Getting Started

### Gradle Dependency

```groovy
dependencies {
    implementation 'io.github.excel-ninja:excelNinja:0.0.1'
}
```

> 📦 No need to add Apache POI manually.  
> ExcelNinja bundles all required dependencies internally.

---

## Usage

### 1. Define a DTO Classes

```java
public class UserReadDto {
    @ExcelReadColumn(headerName = "ID")
    private Long id;

    @ExcelReadColumn(headerName = "Name")
    private String name;

    @ExcelReadColumn(headerName = "Age")
    private Integer age;

    public UserReadDto() {}
}
```

```java
public class UserWriteDto {
    @ExcelWriteColumn(headerName = "ID", order = 0)
    private Long id;

    @ExcelWriteColumn(headerName = "Name", order = 1)
    private String name;

    @ExcelWriteColumn(headerName = "Age", order = 2)
    private Integer age;

    public UserWriteDto(Long id, String name, Integer age) {
        this.id = id;
        this.name = name;
        this.age = age;
    }
}
```

### 2. Read Excel File into DTO List

```java
File file = new File("users.xlsx");
List<UserReadDto> users = NinjaExcel.read(file, UserReadDto.class);
```

### 3. Write DTO List to Excel File

```java
ExcelDocument document = ExcelDocument.createFromEntities(data);
NinjaExcel.write(document, "/path/where/you/want/output.xlsx");
```

---

## Core Components

| Class / Interface      | Description |
|------------------------|-------------|
| `NinjaExcel`           | Main facade for reading/writing Excel |
| `ExcelDocument`        | In-memory Excel representation, supports DTO conversion |
| `ExcelReadColumn`      | Annotation for Excel-to-DTO mapping |
| `ExcelWriteColumn`     | Annotation for DTO-to-Excel mapping |
| `ExcelReader`          | Abstraction for reading Excel files |
| `ExcelWriter`          | Abstraction for writing Excel files |
| `DefaultConverter`     | Type conversion implementation |

---

## Annotations

### `@ExcelReadColumn(headerName = "...")`

Maps a DTO field to a column in Excel when reading.

### `@ExcelWriteColumn(headerName = "...", order = 0)`

Maps a DTO field to a column when writing, with optional column order.
if `order` is not specified, the order is determined by field declaration order.

---

## Logging

NinjaExcel prints an ASCII logo to the console when the class is loaded.  
(Only visible if your logging system displays `INFO` level logs.)

---

## Exception Handling

All checked exceptions (IO, reflection, etc.) are wrapped in runtime exceptions  
so you can use the API without repetitive try-catch blocks.
