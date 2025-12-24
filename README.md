# OfficeHelperOpenxmlVsto

A hybrid .NET library for reading and writing Microsoft Office files with a unique dual-architecture approach:

- **Read**: OpenXML SDK (fast, server-side, no Office required)
- **Write**: VSTO/COM (full feature support, format fidelity)

## 🎯 Features

### Reading (OpenXML SDK)
- ✅ Read PowerPoint (.pptx) files without Microsoft Office installed
- ✅ Read Excel (.xlsx) files without Microsoft Office installed
- ✅ Extract text, shapes, images, tables, and formatting
- ✅ Export to JSON for analysis and processing
- ✅ Cross-platform compatible (Linux, macOS, Windows)

### Writing (VSTO/COM)
- ✅ Create PowerPoint files from JSON data
- ✅ Use templates to preserve master styles and layouts
- ✅ Support for all PowerPoint features (animations, transitions, etc.)
- ✅ Full format fidelity - perfect style preservation
- ✅ Automatic COM object management and cleanup

### Supported Features
- Text boxes and WordArt
- AutoShapes (rectangles, circles, etc.)
- Pictures and images
- Tables
- Groups and connectors
- Fill styles (solid, gradient, pattern, picture)
- Line styles and borders
- Shadow effects
- Text formatting (font, size, color, effects)
- Theme colors and color transforms
- Slide masters and layouts

## 📋 Requirements

### Development Environment
- **Visual Studio** 2019 or 2022
- **.NET Framework** 4.8 SDK
- **Microsoft Office** 2016 or later (with PowerPoint) - required for write operations
- **Windows** operating system (required for COM/VSTO operations)

### NuGet Packages
- `DocumentFormat.OpenXml` 3.3.0
- `Microsoft.Office.Interop.PowerPoint` 15.0.4420.1017
- `Microsoft.Office.Interop.Excel` 15.0.4420.1017
- `Newtonsoft.Json` 13.0.4

## 🚀 Quick Start

### Installation

```bash
git clone https://github.com/Galaxy-Ding/OfficeHelperOpenxmlVsto.git
cd OfficeHelperOpenxmlVsto/OfficeHelperOpenxmVsto
```

**Important**: This project cannot be built with `dotnet build` due to COM references. You must use Visual Studio MSBuild.

### Building with Visual Studio

1. Open `OfficeHelperOpenxmVsto.sln` in Visual Studio 2019/2022
2. Right-click the project → Add → Reference → COM
3. Add the following COM references:
   - ✅ Microsoft Office 16.0 Object Library
   - ✅ Microsoft PowerPoint 16.0 Object Library
4. Build → Build Solution (Ctrl+Shift+B)

### Building with MSBuild (Command Line)

```powershell
# Using Developer Command Prompt
msbuild OfficeHelperOpenXml.csproj /t:Build /p:Configuration=Debug

# Or use the provided PowerShell script
.\FIX_COM_REFERENCE.ps1
```

## 📖 Usage Examples

### Reading PowerPoint to JSON

```csharp
using OfficeHelperOpenXml.Api;

// Read PowerPoint file and export to JSON
string pptxPath = "presentation.pptx";
string outputPath = "output.json";

bool success = OfficeHelperWrapper.AnalyzePowerPointToFile(
    pptxPath, outputPath);

if (success)
{
    Console.WriteLine("Analysis complete!");
}
```

### Creating PowerPoint from JSON

```csharp
using OfficeHelperOpenXml.Api;

// Create PowerPoint from JSON using a template
string templatePath = "template.pptx";
string jsonPath = "data.json";
string outputPath = "output.pptx";

string jsonData = File.ReadAllText(jsonPath);

bool success = OfficeHelperWrapper.WritePowerPointFromJson(
    templatePath, jsonData, outputPath);

if (success)
{
    Console.WriteLine("PowerPoint created successfully!");
}
```

### Using Command Line

```bash
# Extract PowerPoint to JSON
OfficeHelperOpenXml.exe --mode extract --input presentation.pptx --output output.json

# Create PowerPoint from JSON (requires template)
OfficeHelperOpenXml.exe --mode create --input data.json --output output.pptx --template template.pptx

# Auto-detect mode (by file extension)
OfficeHelperOpenXml.exe presentation.pptx output.json
OfficeHelperOpenXml.exe data.json output.pptx template.pptx
```

### Advanced: Using the API Directly

```csharp
using OfficeHelperOpenXml.Api.PowerPoint;
using OfficeHelperOpenXml.Models.Json;

// Create a writer from factory
using (var writer = PowerPointWriterFactory.CreateWriter())
{
    // 1. Open from template
    writer.OpenFromTemplate("template.pptx");
    
    // 2. Clear existing content slides (optional)
    writer.ClearAllContentSlides();
    
    // 3. Prepare JSON data
    var jsonData = new PresentationJsonData
    {
        ContentSlides = new List<SlideJsonData>
        {
            new SlideJsonData
            {
                PageNumber = 1,
                Title = "New Slide",
                Shapes = new List<ShapeJsonData>
                {
                    new ShapeJsonData
                    {
                        Type = "textbox",
                        Name = "Title1",
                        Box = "2,2,20,3",
                        HasText = 1,
                        Text = new List<TextRunJsonData>
                        {
                            new TextRunJsonData
                            {
                                Content = "Hello World!",
                                Font = "Arial",
                                FontSize = 24,
                                FontColor = "RGB(0,0,0)"
                            }
                        }
                    }
                }
            }
        }
    };
    
    // 4. Write data
    writer.WriteFromJsonData(jsonData);
    
    // 5. Save to file
    writer.SaveAs("output.pptx");
}
```

## 📁 Project Structure

```
OfficeHelperOpenxmlVsto/
├── OfficeHelperOpenxmVsto/          # Main project
│   ├── Api/                         # Public API layer
│   │   ├── PowerPoint/              # PowerPoint writer API
│   │   ├── Excel/                   # Excel reader/writer API
│   │   ├── Word/                    # Word reader/writer API
│   │   └── OfficeHelperWrapper.cs  # Static wrapper methods
│   ├── Core/                        # Core implementation
│   │   ├── Readers/                 # OpenXML readers
│   │   ├── Writers/                 # VSTO writers
│   │   ├── Converters/              # JSON converters
│   │   └── Models/                  # Data models
│   ├── Elements/                    # Shape elements
│   ├── Components/                  # Style components
│   ├── Models/Json/                 # JSON data models
│   ├── Utils/                       # Utilities
│   └── Program.cs                   # Console application entry point
├── OfficeHelperOpenxmVsto.Test/     # Test project
└── Documentation files
```

## 🔧 Architecture

### Why This Dual Architecture?

| Operation | Technology | Pros | Cons |
|-----------|-----------|------|------|
| **Reading** | OpenXML SDK | • Fast<br>• No Office required<br>• Cross-platform<br>• Server-safe | • Limited feature support<br>• Complex API |
| **Writing** | VSTO/COM | • Full feature support<br>• Perfect format fidelity<br>• Template support<br>• Simple API | • Requires PowerPoint<br>• Windows only<br>• Slower (starts app) |

This hybrid approach gives you the best of both worlds:
- **Fast, server-side reading** for analysis and extraction
- **Complete, faithful writing** for generation and creation

### Workflow

```
┌─────────────────┐
│  Template PPTX  │
└────────┬────────┘
         │
         ▼
┌─────────────────┐     ┌─────────────────┐
│   Read with     │────▶│    JSON Data    │
│  OpenXML SDK    │     │   (Analysis)    │
└─────────────────┘     └────────┬────────┘
                                 │
                                 ▼
                        ┌─────────────────┐
                        │  Modify/Edit    │
                        │  JSON Data      │
                        └────────┬────────┘
                                 │
                                 ▼
┌─────────────────┐     ┌─────────────────┐
│   Write with    │◀────│  JSON Data      │
│     VSTO/COM    │     │   (Finalized)   │
└────────┬────────┘     └─────────────────┘
         │
         ▼
┌─────────────────┐
│  Output PPTX    │
└─────────────────┘
```

## 📚 API Documentation

- [API Documentation](OfficeHelperOpenxmVsto/API_DOCUMENTATION.md) - Complete API reference
- [Architecture Confirmation](OfficeHelperOpenxmVsto/ARCHITECTURE_CONFIRMATION.md) - Architecture details
- [Build Instructions](OfficeHelperOpenxmVsto/BUILD_INSTRUCTIONS.md) - Build guide
- [VSTO Quick Start](OfficeHelperOpenxmVsto/VSTO_QUICK_START.md) - VSTO development guide

## 🧪 Testing

The project includes comprehensive test coverage:

```bash
# Open test project in Visual Studio
OfficeHelperOpenxmVsto.Test/OfficeHelperOpenXmlVsto.Test.csproj

# Run tests in Visual Studio
Test → Run All Tests
```

Test categories:
- PowerPoint writer integration tests
- Text component property tests
- Color and formatting tests
- Template analysis tests
- Edge case tests

## ⚠️ Known Limitations

1. **Write operations require PowerPoint** - VSTO/COM needs PowerPoint installed on Windows
2. **Cannot use `dotnet build`** - Must use Visual Studio MSBuild due to COM references
3. **Not cross-platform for writing** - Write operations only work on Windows
4. **Performance** - Large files with 100+ shapes may be slow during write operations
5. **File locking** - Cannot write to files that are open in PowerPoint

## 🐛 Troubleshooting

### Build Errors

**Error**: `.NET Core version of MSBuild does not support 'ResolveComReference'`

**Solution**: Use Visual Studio MSBuild instead of `dotnet build`:
```powershell
msbuild OfficeHelperOpenXml.csproj /t:Build /p:Configuration=Debug
```

**Error**: `Cannot get type library for "91493440-5a91-11cf-8700-00aa0060263b"`

**Solution**: 
1. Install Microsoft PowerPoint
2. Add COM references in Visual Studio
3. See [COM_REFERENCE_FIX.md](OfficeHelperOpenxmVsto/COM_REFERENCE_FIX.md)

### Runtime Errors

**Error**: `PowerPoint is not available`

**Solution**: Ensure Microsoft PowerPoint is installed and can run.

**Error**: `Template file not found`

**Solution**: Check file paths are correct and use absolute paths.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the GNU General Public License v2.0 - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **DocumentFormat.OpenXml** - OpenXML SDK for reading Office files
- **Microsoft Office Interop** - COM interop for Office automation
- **Newtonsoft.Json** - JSON serialization and deserialization

## 📞 Contact

- Repository: https://github.com/Galaxy-Ding/OfficeHelperOpenxmlVsto
- Issues: https://github.com/Galaxy-Ding/OfficeHelperOpenxmlVsto/issues

---

**Note**: This is an active project. Features may be added or changed. Please check the documentation for the latest information.
