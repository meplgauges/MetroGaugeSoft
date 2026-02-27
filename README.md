# MetroGaugeSoft

MetroGaugeSoft is a .NET 8 WPF-based industrial measurement software with PLC integration, serial communication, database storage, 
reporting, and graphical visualization.

---

## 🔧 Project Information

- Project Type: Windows Desktop Application (WPF)
- Target Framework: net8.0-windows10.0.20348.0
- Platform: AnyCPU
- IDE: Visual Studio 2022 or later

---

## 🛠 Technologies & Libraries

### Core
- .NET 8
- WPF

### Database
- SQL Server Express
- Microsoft.Data.SqlClient (6.1.1)

### PLC Communication
- Mitsubishi MX Component (ActUtlType64Lib v5.0)
- PMcProtocol (2.2.0)

### Serial Communication
- System.IO.Ports (9.0.8)

### Charts & Visualization
- LiveChartsCore.SkiaSharpView.WPF (2.0.0-rc6.1)
- ScottPlot.WPF (4.1.67)
- SkiaSharp (3.119.2-preview.1)

### Reporting & Export
- ClosedXML (0.105.0)
- DocumentFormat.OpenXml (3.3.0)
- PdfSharpCore (1.3.67)

### UI Enhancements
- WPF-UI (4.0.3)
- WPF-UI.DependencyInjection (4.0.3)
- Extended.Wpf.Toolkit (5.0.0)
- LoadingSpinner.WPF (1.0.0)

---

## 🗄 Database Configuration

Database Name:
MetroGaugeSoft

Default Connection String:

Data Source=DESKTOP-1HO9NHN\SQLEXPRESS01;
Initial Catalog=MetroGaugeSoft;
Integrated Security=True;
TrustServerCertificate=True;

⚠ IMPORTANT:
Change the SQL Server instance name according to your local system.

Example:
Data Source=YOURPC\SQLEXPRESS;

---

## 🧩 Required External Components

1. Mitsubishi MX Component (64-bit)  
   Required for ActUtlType64Lib (PLC communication)

2. Required Local DLL Files:
   - Interop.OrbitCOM.dll
   - OrbitLibrary.dll

If these are missing, the application will not build or run.

---

## 🚀 How To Run The Project

1. Install .NET 8 SDK (Windows Desktop)
   Verify:
   dotnet --version

2. Install SQL Server Express + SSMS  
   Create database:
   MetroGaugeSoft

3. Update connection string in:
   - appsettings.json OR
   - App.config

4. Install Mitsubishi MX Component (64-bit)

5. Restore NuGet Packages:
   dotnet restore

6. Build Project:
   dotnet build

7. Run Application:
   dotnet run
   OR press F5 in Visual Studio

---

## ⚠ Common Issues

SQL Connection Failed  
→ Check SQL Server instance name.

COM Component Error  
→ Install Mitsubishi MX Component.

Missing DLL Error  
→ Ensure required DLL files exist.

NuGet Package Missing  
→ Run dotnet restore.

---

## 📌 Notes

- Ensure SQL Server service is running.
- Use Release build for production.
- Avoid hardcoding connection strings in production systems.

---

Industrial Measurement & PLC Integrated Monitoring System
