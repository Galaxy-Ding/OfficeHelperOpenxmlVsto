# OfficeHelperOpenxmlVsto

一个独特的双架构 .NET 库，用于读写 Microsoft Office 文件：

- **读取**: OpenXML SDK（快速、服务器端、无需安装 Office）
- **写入**: VSTO/COM（完整功能支持、格式保真）

## 🎯 功能特性

### 读取功能（OpenXML SDK）
- ✅ 无需安装 Microsoft Office 即可读取 PowerPoint (.pptx) 文件
- ✅ 无需安装 Microsoft Office 即可读取 Excel (.xlsx) 文件
- ✅ 提取文本、形状、图片、表格和格式
- ✅ 导出为 JSON 用于分析和处理
- ✅ 跨平台兼容（Linux、macOS、Windows）

### 写入功能（VSTO/COM）
- ✅ 从 JSON 数据创建 PowerPoint 文件
- ✅ 使用模板保持母版样式和布局
- ✅ 支持所有 PowerPoint 功能（动画、过渡效果等）
- ✅ 完整的格式保真 - 完美的样式保留
- ✅ 自动 COM 对象管理和清理

### 支持的功能
- 文本框和艺术字
- 自选形状（矩形、圆形等）
- 图片和图像
- 表格
- 组合和连接符
- 填充样式（纯色、渐变、图案、图片）
- 线条样式和边框
- 阴影效果
- 文本格式（字体、大小、颜色、效果）
- 主题颜色和颜色转换
- 幻灯片母版和版式

## 📋 系统要求

### 开发环境
- **Visual Studio** 2019 或 2022
- **.NET Framework** 4.8 SDK
- **Microsoft Office** 2016 或更高版本（包含 PowerPoint）- 写入操作必需
- **Windows** 操作系统（COM/VSTO 操作必需）

### NuGet 包
- `DocumentFormat.OpenXml` 3.3.0
- `Microsoft.Office.Interop.PowerPoint` 15.0.4420.1017
- `Microsoft.Office.Interop.Excel` 15.0.4420.1017
- `Newtonsoft.Json` 13.0.4

## 🚀 快速开始

### 安装

```bash
git clone https://github.com/Galaxy-Ding/OfficeHelperOpenxmlVsto.git
cd OfficeHelperOpenxmlVsto/OfficeHelperOpenxmVsto
```

**重要提示**: 由于 COM 引用，此项目无法使用 `dotnet build` 构建。必须使用 Visual Studio MSBuild。

### 使用 Visual Studio 构建

1. 在 Visual Studio 2019/2022 中打开 `OfficeHelperOpenxmVsto.sln`
2. 右键点击项目 → 添加 → 引用 → COM
3. 添加以下 COM 引用：
   - ✅ Microsoft Office 16.0 Object Library
   - ✅ Microsoft PowerPoint 16.0 Object Library
4. 生成 → 生成解决方案（Ctrl+Shift+B）

### 使用 MSBuild 构建（命令行）

```powershell
# 使用开发人员命令提示符
msbuild OfficeHelperOpenXml.csproj /t:Build /p:Configuration=Debug

# 或使用提供的 PowerShell 脚本
.\FIX_COM_REFERENCE.ps1
```

## 📖 使用示例

### 读取 PowerPoint 导出为 JSON

```csharp
using OfficeHelperOpenXml.Api;

// 读取 PowerPoint 文件并导出为 JSON
string pptxPath = "presentation.pptx";
string outputPath = "output.json";

bool success = OfficeHelperWrapper.AnalyzePowerPointToFile(
    pptxPath, outputPath);

if (success)
{
    Console.WriteLine("分析完成！");
}
```

### 从 JSON 创建 PowerPoint

```csharp
using OfficeHelperOpenXml.Api;

// 使用模板从 JSON 创建 PowerPoint
string templatePath = "template.pptx";
string jsonPath = "data.json";
string outputPath = "output.pptx";

string jsonData = File.ReadAllText(jsonPath);

bool success = OfficeHelperWrapper.WritePowerPointFromJson(
    templatePath, jsonData, outputPath);

if (success)
{
    Console.WriteLine("PowerPoint 创建成功！");
}
```

### 使用命令行

```bash
# 提取 PowerPoint 为 JSON
OfficeHelperOpenXml.exe --mode extract --input presentation.pptx --output output.json

# 从 JSON 创建 PowerPoint（需要模板）
OfficeHelperOpenXml.exe --mode create --input data.json --output output.pptx --template template.pptx

# 自动检测模式（通过文件扩展名）
OfficeHelperOpenXml.exe presentation.pptx output.json
OfficeHelperOpenXml.exe data.json output.pptx template.pptx
```

### 高级用法：直接使用 API

```csharp
using OfficeHelperOpenXml.Api.PowerPoint;
using OfficeHelperOpenXml.Models.Json;

// 从工厂创建写入器
using (var writer = PowerPointWriterFactory.CreateWriter())
{
    // 1. 从模板打开
    writer.OpenFromTemplate("template.pptx");
    
    // 2. 清除现有内容幻灯片（可选）
    writer.ClearAllContentSlides();
    
    // 3. 准备 JSON 数据
    var jsonData = new PresentationJsonData
    {
        ContentSlides = new List<SlideJsonData>
        {
            new SlideJsonData
            {
                PageNumber = 1,
                Title = "新幻灯片",
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
                                Content = "你好，世界！",
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
    
    // 4. 写入数据
    writer.WriteFromJsonData(jsonData);
    
    // 5. 保存到文件
    writer.SaveAs("output.pptx");
}
```

## 📁 项目结构

```
OfficeHelperOpenxmlVsto/
├── OfficeHelperOpenxmVsto/          # 主项目
│   ├── Api/                         # 公共 API 层
│   │   ├── PowerPoint/              # PowerPoint 写入器 API
│   │   ├── Excel/                   # Excel 读取器/写入器 API
│   │   ├── Word/                    # Word 读取器/写入器 API
│   │   └── OfficeHelperWrapper.cs  # 静态包装方法
│   ├── Core/                        # 核心实现
│   │   ├── Readers/                 # OpenXML 读取器
│   │   ├── Writers/                 # VSTO 写入器
│   │   ├── Converters/              # JSON 转换器
│   │   └── Models/                  # 数据模型
│   ├── Elements/                    # 形状元素
│   ├── Components/                  # 样式组件
│   ├── Models/Json/                 # JSON 数据模型
│   ├── Utils/                       # 工具类
│   └── Program.cs                   # 控制台应用程序入口
├── OfficeHelperOpenxmVsto.Test/     # 测试项目
└── 文档文件
```

## 🔧 架构设计

### 为什么采用双架构？

| 操作 | 技术 | 优点 | 缺点 |
|-----------|-----------|------|------|
| **读取** | OpenXML SDK | • 快速<br>• 无需 Office<br>• 跨平台<br>• 服务器安全 | • 功能支持有限<br>• API 复杂 |
| **写入** | VSTO/COM | • 完整功能支持<br>• 完美格式保真<br>• 模板支持<br>• 简单 API | • 需要 PowerPoint<br>• 仅限 Windows<br>• 较慢（启动应用） |

这种混合架构为您提供了两全其美的优势：
- **快速的服务器端读取** 用于分析和提取
- **完整的保真写入** 用于生成和创建

### 工作流程

```
┌─────────────────┐
│  模板 PPTX      │
└────────┬────────┘
         │
         ▼
┌─────────────────┐     ┌─────────────────┐
│   使用          │────▶│    JSON 数据    │
│  OpenXML SDK    │     │   (分析)        │
│     读取        │     └────────┬────────┘
└─────────────────┘             │
                               │
                               ▼
                      ┌─────────────────┐
                      │  修改/编辑      │
                      │  JSON 数据      │
                      └────────┬────────┘
                               │
                               ▼
┌─────────────────┐     ┌─────────────────┐
│   使用          │◀────│  JSON 数据      │
│  VSTO/COM       │     │   (最终版)      │
│     写入        │     └─────────────────┘
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  输出 PPTX      │
└─────────────────┘
```

## 📚 API 文档

- [API 文档](OfficeHelperOpenxmVsto/API_DOCUMENTATION.md) - 完整的 API 参考
- [架构确认](OfficeHelperOpenxmVsto/ARCHITECTURE_CONFIRMATION.md) - 架构详情
- [构建说明](OfficeHelperOpenxmVsto/BUILD_INSTRUCTIONS.md) - 构建指南
- [VSTO 快速入门](OfficeHelperOpenxmVsto/VSTO_QUICK_START.md) - VSTO 开发指南

## 🧪 测试

项目包含全面的测试覆盖：

```bash
# 在 Visual Studio 中打开测试项目
OfficeHelperOpenxmVsto.Test/OfficeHelperOpenXmlVsto.Test.csproj

# 在 Visual Studio 中运行测试
测试 → 运行所有测试
```

测试类别：
- PowerPoint 写入器集成测试
- 文本组件属性测试
- 颜色和格式测试
- 模板分析测试
- 边界情况测试

## ⚠️ 已知限制

1. **写入操作需要 PowerPoint** - VSTO/COM 需要在 Windows 上安装 PowerPoint
2. **无法使用 `dotnet build`** - 由于 COM 引用必须使用 Visual Studio MSBuild
3. **写入不支持跨平台** - 写入操作仅适用于 Windows
4. **性能** - 包含 100+ 个形状的大型文件在写入操作期间可能较慢
5. **文件锁定** - 无法写入在 PowerPoint 中打开的文件

## 🐛 故障排除

### 构建错误

**错误**: `.NET Core 版本的 MSBuild 不支持 'ResolveComReference'`

**解决方案**: 使用 Visual Studio MSBuild 而不是 `dotnet build`：
```powershell
msbuild OfficeHelperOpenXml.csproj /t:Build /p:Configuration=Debug
```

**错误**: `无法获取 "91493440-5a91-11cf-8700-00aa0060263b" 的类型库`

**解决方案**: 
1. 安装 Microsoft PowerPoint
2. 在 Visual Studio 中添加 COM 引用
3. 查看 [COM_REFERENCE_FIX.md](OfficeHelperOpenxmVsto/COM_REFERENCE_FIX.md)

### 运行时错误

**错误**: `PowerPoint 不可用`

**解决方案**: 确保已安装 Microsoft PowerPoint 并且能够运行。

**错误**: `找不到模板文件`

**解决方案**: 检查文件路径是否正确，并使用绝对路径。

## 🤝 贡献

欢迎贡献！请随时提交 Pull Request。

1. Fork 本仓库
2. 创建您的功能分支（`git checkout -b feature/AmazingFeature`）
3. 提交您的更改（`git commit -m 'Add some AmazingFeature'`）
4. 推送到分支（`git push origin feature/AmazingFeature`）
5. 开启 Pull Request

## 📄 许可证

本项目在 GNU General Public License v2.0 下许可 - 详见 [LICENSE](LICENSE) 文件。

## 🙏 致谢

- **DocumentFormat.OpenXml** - 用于读取 Office 文件的 OpenXML SDK
- **Microsoft Office Interop** - 用于 Office 自动化的 COM 互操作
- **Newtonsoft.Json** - JSON 序列化和反序列化

## 📞 联系方式

- 仓库: https://github.com/Galaxy-Ding/OfficeHelperOpenxmlVsto
- 问题: https://github.com/Galaxy-Ding/OfficeHelperOpenxmlVsto/issues

---

**注意**: 这是一个活跃的项目。功能可能会添加或更改。请查看文档以获取最新信息。
