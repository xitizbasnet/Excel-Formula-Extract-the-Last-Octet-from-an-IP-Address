# Excel Formula: Extract the Last Octet from an IP Address

> **Applies to:** Microsoft Excel 365, Excel 2021 or later

Extract the **last octet** from an IPv4 address and return it as a **numeric value** using a simple Excel formula.

---

## 📖 Overview

This formula extracts the final number (last octet) from an IPv4 address and converts it into an integer. It is useful for network administrators, IT professionals, and anyone working with IP address inventories in Microsoft Excel.

---

## 🎯 Problem

Given IPv4 addresses such as:

```text
192.168.100.20
10.12.9.141
172.16.17.54
```

Extract only the last octet:

```text
20
141
54
```

---

## ✅ Solution

Use the following Excel formula:

```excel
=INT(VALUE(TEXTAFTER(H2,".",-1)))
```

> **Note**
>
> Replace **H2** with the cell containing your IP address.

---

## ⚙️ How It Works

The formula consists of three functions:

| Function | Description |
|----------|-------------|
| `TEXTAFTER(H2,".",-1)` | Returns the text that appears after the last period (`.`) in the IP address. |
| `VALUE(...)` | Converts the extracted text into a numeric value. |
| `INT(...)` | Returns the result as an integer, removing any decimal portion if present. |

---

## 📝 Example

| IP Address | Result |
|------------|-------:|
| 192.168.100.20 | 20 |
| 10.12.9.141 | 141 |
| 172.16.17.54 | 54 |

---

## 🔄 Alternative Formula

You can also use the following shorter version:

```excel
=--TEXTAFTER(H2,".",-1)
```

### How It Works

The **double-unary (`--`)** operator converts the extracted text directly into a numeric value without using the `VALUE()` function.

---

## 📋 Requirements

- Microsoft Excel 365
- Microsoft Excel 2021 or later
- Support for the `TEXTAFTER()` function

---

## 💡 Use Cases

This formula is useful for:

- Network inventory management
- IP address sorting
- Asset management
- Device inventory reporting
- Network documentation
- Log analysis
- IP-based reporting
- Infrastructure auditing

---

## ⚠️ Notes

- This formula is designed for **IPv4 addresses**.
- Ensure the source cell contains a valid IP address.
- Invalid or improperly formatted values may return an error.
- The result is returned as a **numeric value**, allowing it to be used in sorting, filtering, calculations, and conditional formatting.

---

## 🔍 Formula Breakdown

```text
TEXTAFTER(H2,".",-1)
        │
        └── Extracts everything after the last "."
                 │
                 ▼
              "141"
                 │
          VALUE(...)
                 │
                 ▼
               141
                 │
            INT(...)
                 │
                 ▼
               141
```

---

## 📚 Related Topics

- Excel Text Functions
- `TEXTAFTER()`
- `VALUE()`
- `INT()`
- IPv4 Addressing

---

## 👤 Author

**Xitiz Basnet**

---

⭐ **If this formula saved you time, consider starring the repository.**
