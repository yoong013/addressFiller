# AddressFilter

**AddressFilter** is a lightweight Windows Forms application that lets you import customer address data from Excel, edit it in a grid, automatically map city and state fields based on postcode, and then export the completed dataset back to Excel.

---

## Features

1. **Import Excel** (*.xlsx, *.xls) without treating the first row as headers—loads all rows into a preset 13‑column table:
   - `custID`, `CustName`, `UnitNumber`, `Address1`, `Address2`, `City`, `Postcode`, `State`, `Col9`, `Contact`, `Col11`, `Col12`, `Col13`
2. **Editable Grid**: correct any data in‑place (names, addresses, postcodes).
3. **Map City & State**: click **Map** to read the postcode (column 7), look up city and state from `postcode_city.json`, and write them into column 6 (`City`) and column 8 (`State`).
4. **Export Back**: click **Export** to save the edited and mapped table to a new Excel workbook.

---

## Prerequisites

- Windows 10 or later
- [.NET Framework 4.8](https://dotnet.microsoft.com/download/dotnet-framework/net48) or [.NET 6 SDK](https://dotnet.microsoft.com/download)
- Visual Studio 2022 (Community or higher) with **Desktop development with C#** workload (optional)
- `postcode_city.json` file (Malaysia postcode→city/state map)

---

## Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/your-username/AddressFilter.git
cd AddressFilter
```

### 2. Prepare the postcode map

Place **`postcode_city.json`** in the project root (or next to the compiled EXE). It must contain a JSON array of states, each with cities and postcodes, for example:

```json
[
  {
    "name": "Selangor",
    "city": [
      { "name": "Petaling Jaya", "postcode": ["46000","46100"] },
      { "name": "Shah Alam", "postcode": ["40000","40100"] }
    ]
  },
  …
]
```

### 3. Build the project

- **In Visual Studio**: open `AddressFilter.sln`, set target framework to .NET 4.8 (or .NET 6), build.
- **Using CLI**:
  ```bash
  dotnet build AddressFilter.csproj -c Release
  ```

### 4. Run the application

- **From Visual Studio**: press **F5** or **Ctrl+F5**.
- **From published EXE**:
  1. `dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true`
  2. Copy `AddressFilter.exe` and `postcode_city.json` into one folder and double-click the EXE.

---

## Usage

1. **Import Excel**: click **Import…** and select your `.xlsx`/`.xls` file. All rows—including header row—are loaded into the grid under the fixed columns.
2. **Edit Data**: click any cell to correct or override values (e.g. fix typos before mapping).
3. **Map**: click **Map** to fill the **City** and **State** columns based on the **Postcode** column. Unmatched postcodes remain blank.
4. **Export**: click **Export…**, choose a save location, and the app writes out the current grid to a new Excel file.

---

## Customization

- To change column presets or positions, open **Form1.cs** and adjust the `dt.Columns.Add(...)` calls in `btnImport_Click`.
- To support different JSON formats (e.g. flat `Dictionary<string,string>`), modify the JSON‑loading logic in the constructor of **Form1**.
- You can add progress indicators, column‑mapping dialogs, or error logs as needed.

---

## Troubleshooting

- **Permission errors** adding `.vs/` files: a sample `.gitignore` is included to ignore Visual Studio artifacts.
- **Line-ending warnings**: if you see `LF will be replaced by CRLF`, add a `.gitattributes` with `*.json text eol=lf` to enforce LF.
- **Mapping errors**: ensure your JSON keys (postcodes) match exactly the text in the **Postcode** column.

---

## License

This project is released under the [MIT License](LICENSE). Feel free to fork and contribute!

