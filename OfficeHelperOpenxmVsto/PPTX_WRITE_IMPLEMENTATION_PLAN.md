# PPTX 写入策略实施计划

## 📋 概述

本文档提供各个写入策略的详细实施计划，包括代码示例、步骤说明和测试方案。

---

## 🎯 策略1：改进的 VSTO（保存和恢复 DisplayAlerts）

### 目标
快速修复 DisplayAlerts 全局设置问题，确保不影响用户的其他 PPTX 文件。

### 实施步骤

#### 步骤1：添加字段保存原始值

**文件：** `Api/PowerPoint/PowerPointWriter.cs`

**修改位置：** 类字段声明区域（约第22行）

```csharp
public class PowerPointWriter : IPowerPointWriter
{
    private Application _app;
    private Presentation _presentation;
    private VstoSlideWriter _slideWriter;
    private JsonToVstoConverter _converter;
    private bool _disposed = false;
    private bool _appCreatedByUs = false;
    
    // ⭐ 新增字段
    private PpAlertLevel _originalDisplayAlerts = PpAlertLevel.ppAlertsAll;
    private bool _displayAlertsModified = false;
}
```

#### 步骤2：在 OpenFromTemplate 中保存原始值

**文件：** `Api/PowerPoint/PowerPointWriter.cs`

**修改位置：** `OpenFromTemplate()` 方法（约第98行）

```csharp
public bool OpenFromTemplate(string templatePath)
{
    // ... 前面的代码保持不变 ...
    
    try
    {
        // ⭐ 策略1：智能实例管理 - 尝试获取现有的 PowerPoint 实例
        try
        {
            _app = (Application)Marshal.GetActiveObject("PowerPoint.Application");
            _appCreatedByUs = false;
            logger.LogInfo("已连接到现有的 PowerPoint 实例");
        }
        catch (COMException)
        {
            _app = new Application();
            _appCreatedByUs = true;
            logger.LogInfo("创建了新的 PowerPoint 实例");
            
            try
            {
                _app.Visible = MsoTriState.msoFalse;
            }
            catch (COMException) { }
        }
        
        // ⭐ 新增：保存原始 DisplayAlerts 值
        try
        {
            _originalDisplayAlerts = _app.DisplayAlerts;
            _app.DisplayAlerts = PpAlertLevel.ppAlertsNone;
            _displayAlertsModified = true;
            logger.LogInfo($"DisplayAlerts 已设置为 ppAlertsNone（原始值: {_originalDisplayAlerts}）");
        }
        catch (Exception ex)
        {
            logger.LogWarning($"保存 DisplayAlerts 时出错: {ex.Message}");
            // 继续执行，不中断流程
        }
        
        // ... 后面的代码保持不变 ...
    }
    catch (Exception ex)
    {
        // ⭐ 新增：如果出错，尝试恢复 DisplayAlerts
        RestoreDisplayAlerts();
        throw;
    }
}
```

#### 步骤3：添加恢复 DisplayAlerts 的辅助方法

**文件：** `Api/PowerPoint/PowerPointWriter.cs`

**修改位置：** `Cleanup()` 方法之前（约第567行）

```csharp
/// <summary>
/// 恢复 DisplayAlerts 原始值
/// </summary>
private void RestoreDisplayAlerts()
{
    if (!_displayAlertsModified || _app == null)
        return;
    
    try
    {
        _app.DisplayAlerts = _originalDisplayAlerts;
        _displayAlertsModified = false;
        var logger = new Logger();
        logger.LogInfo($"DisplayAlerts 已恢复为原始值: {_originalDisplayAlerts}");
    }
    catch (Exception ex)
    {
        var logger = new Logger();
        logger.LogWarning($"恢复 DisplayAlerts 时出错: {ex.Message}");
        // 不抛出异常，确保清理流程继续
    }
}
```

#### 步骤4：在 Cleanup 中恢复原始值

**文件：** `Api/PowerPoint/PowerPointWriter.cs`

**修改位置：** `Cleanup()` 方法（约第568行）

```csharp
private void Cleanup()
{
    var logger = new Logger();
    try
    {
        logger.LogInfo("[Cleanup] 开始清理资源");
        
        // ⭐ 新增：先恢复 DisplayAlerts，再关闭演示文稿
        RestoreDisplayAlerts();
        
        Close();

        if (_app != null)
        {
            if (_appCreatedByUs)
            {
                // ... 检查演示文稿数量的代码保持不变 ...
            }
            else
            {
                logger.LogInfo("[Cleanup] PowerPoint 实例不是我们创建的，不关闭应用程序");
            }
            
            VstoHelper.ReleaseComObject(_app);
            logger.LogInfo("[Cleanup] PowerPoint 应用程序 COM 对象已释放");
            _app = null;
        }

        // 强制垃圾回收以释放 COM 对象
        logger.LogInfo("[Cleanup] 准备强制垃圾回收");
        VstoHelper.ForceGarbageCollection();
        logger.LogInfo("[Cleanup] 垃圾回收完成，资源清理结束");
    }
    catch (Exception ex)
    {
        logger.LogWarning($"清理资源时出错: {ex.Message}");
        // ⭐ 新增：确保即使出错也恢复 DisplayAlerts
        RestoreDisplayAlerts();
    }
}
```

#### 步骤5：在 Close 方法中也恢复（可选，更安全）

**文件：** `Api/PowerPoint/PowerPointWriter.cs`

**修改位置：** `Close()` 方法（约第544行）

```csharp
public void Close()
{
    var logger = new Logger();
    try
    {
        // ⭐ 新增：在关闭前恢复 DisplayAlerts，确保用户的其他文件有正常的保存提示
        RestoreDisplayAlerts();
        
        if (_presentation != null)
        {
            logger.LogInfo("[Close] 准备关闭演示文稿");
            _presentation.Close();
            logger.LogInfo("[Close] _presentation.Close() 调用返回");
            VstoHelper.ReleaseComObject(_presentation);
            logger.LogInfo("[Close] COM 对象已释放");
            _presentation = null;
        }
    }
    catch (Exception ex)
    {
        logger.LogWarning($"关闭演示文稿时出错: {ex.Message}");
        // ⭐ 确保恢复 DisplayAlerts
        RestoreDisplayAlerts();
    }
}
```

### 测试方案

#### 测试1：基本功能测试
1. 打开一个 PPTX 文件（手动）
2. 运行程序处理另一个 PPTX 文件
3. 验证手动打开的文件没有被关闭
4. 验证手动打开的文件可以正常保存

#### 测试2：DisplayAlerts 恢复测试
1. 记录用户当前的 DisplayAlerts 设置
2. 运行程序
3. 验证程序结束后 DisplayAlerts 恢复为原始值

#### 测试3：异常情况测试
1. 在 OpenFromTemplate 中模拟异常
2. 验证 DisplayAlerts 仍然被恢复
3. 验证资源正确清理

#### 测试4：多次运行测试
1. 连续运行程序多次
2. 验证每次都能正确恢复 DisplayAlerts
3. 验证没有资源泄漏

### 预计时间
- 代码修改：30 分钟
- 测试：30 分钟
- 文档更新：30 分钟
- **总计：1.5 小时**

### 风险评估
- **风险等级：** 🟢 低
- **风险点：**
  - DisplayAlerts 恢复失败（已添加异常处理）
  - 多次恢复导致问题（已添加标志位保护）

---

## 🎯 策略2：纯 OpenXML SDK 写入

### 目标
实现纯 OpenXML SDK 写入，无需 PowerPoint 应用程序，提升性能和架构一致性。

### 架构设计

```
OpenXmlPowerPointWriter
├── OpenFromTemplate()      // 复制模板并打开
├── WriteFromJson()         // 写入 JSON 数据
├── SaveAs()               // 保存文件
└── Dispose()              // 清理资源
```

### 实施步骤

#### 阶段1：基础框架（1周）

##### 步骤1：创建 OpenXmlPowerPointWriter 类

**文件：** `Api/PowerPoint/OpenXmlPowerPointWriter.cs`（新建）

```csharp
using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeHelperOpenXml.Core.Converters;
using OfficeHelperOpenXml.Models.Json;
using OfficeHelperOpenXml.Utils;

namespace OfficeHelperOpenXml.Api.PowerPoint
{
    /// <summary>
    /// 基于 OpenXML SDK 的 PowerPoint 写入器
    /// </summary>
    public class OpenXmlPowerPointWriter : IPowerPointWriter
    {
        private PresentationDocument _document;
        private string _tempPath;
        private JsonToOpenXmlConverter _converter;
        private bool _disposed = false;

        public bool OpenFromTemplate(string templatePath)
        {
            var logger = new Logger();
            
            if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
            {
                logger.LogError("模板文件不存在");
                return false;
            }

            try
            {
                // 创建临时文件副本
                _tempPath = Path.Combine(
                    Path.GetTempPath(),
                    $"pptx_temp_{Guid.NewGuid():N}.pptx"
                );
                
                File.Copy(templatePath, _tempPath, true);
                
                // 打开文档（可写模式）
                _document = PresentationDocument.Open(_tempPath, true);
                
                _converter = new JsonToOpenXmlConverter();
                
                logger.LogSuccess($"成功打开模板文件: {templatePath}");
                return true;
            }
            catch (Exception ex)
            {
                logger.LogError($"打开模板文件失败: {ex.Message}");
                Cleanup();
                return false;
            }
        }

        public bool WriteFromJson(string jsonData)
        {
            var logger = new Logger();
            
            if (string.IsNullOrEmpty(jsonData))
            {
                logger.LogError("JSON 数据不能为空");
                return false;
            }

            try
            {
                var presentationData = _converter?.ParseJson(jsonData);
                if (presentationData == null)
                {
                    logger.LogError("JSON 解析失败");
                    return false;
                }

                return WriteFromJsonData(presentationData);
            }
            catch (Exception ex)
            {
                logger.LogError($"从 JSON 写入内容失败: {ex.Message}");
                return false;
            }
        }

        public bool WriteFromJsonData(PresentationJsonData jsonData)
        {
            // TODO: 实现写入逻辑
            return true;
        }

        public bool ClearAllContentSlides()
        {
            // TODO: 实现清除逻辑
            return true;
        }

        public bool SaveAs(string outputPath)
        {
            var logger = new Logger();
            
            if (string.IsNullOrEmpty(outputPath))
            {
                logger.LogError("输出文件路径不能为空");
                return false;
            }

            if (_document == null)
            {
                logger.LogError("演示文稿未打开");
                return false;
            }

            try
            {
                // 确保输出目录存在
                var directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // 保存文档
                _document.Save();
                
                // 复制到目标位置
                File.Copy(_tempPath, outputPath, true);
                
                logger.LogSuccess($"文件已保存: {outputPath}");
                return true;
            }
            catch (Exception ex)
            {
                logger.LogError($"保存文件失败: {ex.Message}");
                return false;
            }
        }

        public void Close()
        {
            if (_document != null)
            {
                _document.Close();
                _document = null;
            }
        }

        private void Cleanup()
        {
            Close();
            
            // 删除临时文件
            if (!string.IsNullOrEmpty(_tempPath) && File.Exists(_tempPath))
            {
                try
                {
                    File.Delete(_tempPath);
                }
                catch { }
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                Cleanup();
                _disposed = true;
            }
        }
    }
}
```

##### 步骤2：创建 JsonToOpenXmlConverter

**文件：** `Core/Converters/JsonToOpenXmlConverter.cs`（新建）

```csharp
using OfficeHelperOpenXml.Models.Json;
using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Core.Converters
{
    /// <summary>
    /// JSON 到 OpenXML 转换器
    /// </summary>
    public class JsonToOpenXmlConverter
    {
        public PresentationJsonData ParseJson(string jsonData)
        {
            try
            {
                return JsonConvert.DeserializeObject<PresentationJsonData>(jsonData);
            }
            catch
            {
                return null;
            }
        }
    }
}
```

##### 步骤3：创建 OpenXmlSlideWriter

**文件：** `Core/Writers/OpenXmlSlideWriter.cs`（新建）

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeHelperOpenXml.Models.Json;

namespace OfficeHelperOpenXml.Core.Writers
{
    /// <summary>
    /// OpenXML 幻灯片写入器
    /// </summary>
    public class OpenXmlSlideWriter
    {
        private PresentationPart _presentationPart;

        public OpenXmlSlideWriter(PresentationPart presentationPart)
        {
            _presentationPart = presentationPart;
        }

        public void WriteSlides(List<SlideJsonData> slidesData)
        {
            // TODO: 实现写入逻辑
        }

        private SlidePart CreateSlide(SlideJsonData slideData)
        {
            // TODO: 创建幻灯片
            return null;
        }
    }
}
```

#### 阶段2：形状支持（1周）

##### 步骤1：实现文本框创建

**文件：** `Core/Writers/OpenXmlShapeWriter.cs`（新建）

```csharp
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using OfficeHelperOpenXml.Models.Json;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Core.Writers
{
    /// <summary>
    /// OpenXML 形状写入器
    /// </summary>
    public class OpenXmlShapeWriter
    {
        public Shape CreateTextBox(ShapeJsonData shapeData)
        {
            var shape = new Shape();
            
            // 设置形状属性
            shape.NonVisualShapeProperties = CreateNonVisualShapeProperties(shapeData);
            shape.ShapeProperties = CreateShapeProperties(shapeData);
            shape.TextBody = CreateTextBody(shapeData);
            
            return shape;
        }

        private NonVisualShapeProperties CreateNonVisualShapeProperties(ShapeJsonData shapeData)
        {
            // TODO: 实现
            return new NonVisualShapeProperties();
        }

        private ShapeProperties CreateShapeProperties(ShapeJsonData shapeData)
        {
            // TODO: 实现
            return new ShapeProperties();
        }

        private TextBody CreateTextBody(ShapeJsonData shapeData)
        {
            // TODO: 实现
            return new TextBody();
        }
    }
}
```

#### 阶段3：格式支持（1周）

实现文本格式、填充、边框、阴影等。

#### 阶段4：高级功能（1周）

实现表格、图片等复杂元素。

### 测试方案

#### 单元测试
- 测试基本写入功能
- 测试各种形状类型
- 测试格式保持
- 测试异常处理

#### 集成测试
- 与现有读取器对比
- 与 VSTO 写入器对比
- 性能测试

### 预计时间
- **总计：4 周**
- 阶段1：1 周
- 阶段2：1 周
- 阶段3：1 周
- 阶段4：1 周

### 风险评估
- **风险等级：** 🟡 中
- **风险点：**
  - OpenXML 结构复杂，可能遗漏某些属性
  - 格式保真度可能不如 VSTO
  - 开发时间长

---

## 🎯 策略4：隔离 VSTO 实例

### 目标
创建完全隔离的 PowerPoint 实例，不影响用户的其他文件。

### 实施步骤

#### 步骤1：移除 Marshal.GetActiveObject 调用

**文件：** `Api/PowerPoint/PowerPointWriter.cs`

**修改位置：** `OpenFromTemplate()` 方法（约第74行）

```csharp
public bool OpenFromTemplate(string templatePath)
{
    // ... 前面的代码保持不变 ...
    
    try
    {
        // ⭐ 修改：始终创建新实例，不获取现有实例
        _app = new Application();
        _appCreatedByUs = true;  // 标记为我们创建的实例
        
        logger.LogInfo("创建了新的 PowerPoint 实例（隔离模式）");
        
        try
        {
            _app.Visible = MsoTriState.msoFalse;
        }
        catch (COMException)
        {
            // 某些版本的 PowerPoint 不允许隐藏窗口，忽略此错误
        }
        
        // 保存和恢复 DisplayAlerts（参考策略1）
        try
        {
            _originalDisplayAlerts = _app.DisplayAlerts;
            _app.DisplayAlerts = PpAlertLevel.ppAlertsNone;
            _displayAlertsModified = true;
        }
        catch (Exception ex)
        {
            logger.LogWarning($"保存 DisplayAlerts 时出错: {ex.Message}");
        }
        
        // ... 后面的代码保持不变 ...
    }
    catch (Exception ex)
    {
        RestoreDisplayAlerts();
        throw;
    }
}
```

### 测试方案

#### 测试1：隔离性测试
1. 打开一个 PPTX 文件（手动）
2. 运行程序处理另一个 PPTX 文件
3. 验证手动打开的文件没有被关闭
4. 验证程序创建了新的 PowerPoint 进程

#### 测试2：资源清理测试
1. 运行程序
2. 验证程序结束后 PowerPoint 进程被正确关闭
3. 验证没有资源泄漏

### 预计时间
- 代码修改：30 分钟
- 测试：30 分钟
- **总计：1 小时**

### 风险评估
- **风险等级：** 🟢 低
- **风险点：**
  - 性能开销（每次创建新进程）
  - 可能创建多个 PowerPoint 进程

---

## 📊 实施时间表

### 方案A：快速修复（策略1）
- **第1天：** 实施策略1（1-2小时）
- **第2天：** 测试验证
- **总计：** 1-2 天

### 方案B：长期优化（策略2）
- **第1-2周：** 基础框架
- **第3周：** 形状支持
- **第4周：** 格式支持
- **第5周：** 高级功能
- **总计：** 4-5 周

### 方案C：混合实施（策略1 + 策略2）
- **第1天：** 实施策略1（快速修复）
- **第2-5周：** 并行开发策略2
- **总计：** 5 周（但第1天就解决问题）

---

## 🎯 决策建议

### 如果时间紧迫（< 1天）：
→ **选择策略1**（改进的 VSTO）

### 如果有1-2周时间：
→ **选择策略1 + 策略2并行**（快速修复 + 长期优化）

### 如果有1个月以上时间：
→ **选择策略2**（纯 OpenXML 写入）

### 如果性能要求不高：
→ **选择策略4**（隔离 VSTO 实例）

---

**文档创建时间：** 2025-01-XX  
**最后更新：** 2025-01-XX  
**状态：** 📋 待实施

