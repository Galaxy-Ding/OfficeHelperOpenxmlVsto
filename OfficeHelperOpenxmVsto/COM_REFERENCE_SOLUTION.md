# COM 引用问题解决方案

## 问题描述

出现以下错误：
1. `无法获取类型库"91493440-5a91-11cf-8700-00aa0060263b"版本 1.8 的文件路径。库没有注册。`
2. `类型"MsoTriState"在未引用的程序集中定义。必须添加对程序集"office, Version=15.0.0.0"的引用。`

## 根本原因

1. **COM 类型库未注册**：Office 的 COM 组件没有正确注册到系统中
2. **版本不匹配**：NuGet 包 `Microsoft.Office.Interop.PowerPoint` (15.0.4420.1017) 需要 `office, Version=15.0.0.0`，但本地的 `Interop\Microsoft.Office.Core.dll` 是 2.8.0.0 版本

## 解决方案

### 方案 1：重新注册 Office COM 组件（推荐）

1. **以管理员身份打开命令提示符**

2. **导航到 Office 安装目录**
   ```cmd
   cd "C:\Program Files\Microsoft Office\Office16"
   ```
   或
   ```cmd
   cd "C:\Program Files (x86)\Microsoft Office\Office16"
   ```

3. **重新注册 Office 组件**
   ```cmd
   regsvr32 /s MSO.DLL
   regsvr32 /s MSPPT.OLB
   ```

4. **在 Visual Studio 中添加 COM 引用**
   - 右键点击项目 → 添加 → 引用
   - 选择 **COM** 选项卡
   - 勾选：
     - ✅ `Microsoft Office 16.0 Object Library`
     - ✅ `Microsoft PowerPoint 16.0 Object Library`
   - 点击确定

5. **清理并重新构建项目**
   ```cmd
   dotnet clean
   dotnet build
   ```

### 方案 2：使用 tlbimp 重新生成 Microsoft.Office.Core.dll

如果方案 1 不起作用，可以手动生成正确版本的互操作程序集：

1. **找到 tlbimp 工具**
   - 通常位于：`C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\tlbimp.exe`
   - 或使用 Visual Studio Developer Command Prompt

2. **生成互操作程序集**
   ```cmd
   tlbimp "C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL" `
          /out:"OfficeHelperOpenxmVsto\Interop\Microsoft.Office.Core.dll" `
          /namespace:Microsoft.Office.Core `
          /asmversion:15.0.0.0 `
          /publickey:71e9bce111e9429c
   ```

3. **更新项目文件引用**
   项目文件已配置为引用 `Interop\Microsoft.Office.Core.dll`，确保生成的文件版本正确即可。

### 方案 3：使用 Office PIA（Primary Interop Assemblies）

如果 Office 已安装 PIA，可以直接引用：

1. **查找 PIA 位置**
   - 通常在：`C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll`

2. **在项目文件中添加引用**
   ```xml
   <Reference Include="office, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c">
     <HintPath>C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll</HintPath>
   </Reference>
   ```

## 当前项目配置

项目已配置为：
- ✅ 使用 NuGet 包 `Microsoft.Office.Interop.PowerPoint` (15.0.4420.1017)
- ✅ 引用本地 `Interop\Microsoft.Office.Core.dll`
- ⚠️ 但版本不匹配（需要 15.0.0.0，当前是 2.8.0.0）

## 推荐操作步骤

1. **首先尝试方案 1**（重新注册 COM 组件并在 Visual Studio 中添加 COM 引用）
2. **如果方案 1 失败，使用方案 2**（重新生成互操作程序集）
3. **如果方案 2 也失败，使用方案 3**（直接引用 Office PIA）

## 验证修复

构建成功后，应该不再看到：
- ❌ `无法获取类型库...库没有注册`
- ❌ `类型"MsoTriState"在未引用的程序集中定义`
- ❌ `命名空间"Microsoft.Office"中不存在类型或命名空间名"Interop"`

## 注意事项

- 确保 Office 已正确安装
- 确保使用管理员权限执行注册命令
- 如果使用 Visual Studio，建议在 Visual Studio 中添加 COM 引用，而不是手动编辑项目文件
- 版本匹配很重要：PowerPoint 15.0 需要 Office 15.0

