# SCOPSR & SHBB 文章保存器 - 测试指南

## 脚本功能概述
- 支持 SCOPSR (scopsr.gov.cn) 和 SHBB (shbb.gov.cn) 两个网站
- 在搜索页面提供批量保存功能
- 在文章详情页提供单篇文章保存功能
- 保存为 Word 兼容的 DOCX 格式文件

## 如何验证脚本是否正常工作

### 1. 脚本加载验证
- 访问支持的网站后，右上角会短暂显示绿色的"SCOPSR/SHBB脚本已加载"提示
- 如果没有看到此提示，说明脚本可能没有正确加载

### 2. 打开浏览器控制台
- 按 F12 或右键选择"检查"打开开发者工具
- 切换到"Console"（控制台）标签
- 正常情况下应该能看到以下类型的调试信息：

```
[SCOPSR-SHBB Script] 脚本开始执行...
[SCOPSR-SHBB Script] 检测到网站: scopsr.gov.cn
[SCOPSR-SHBB Script] 当前URL: https://www.scopsr.gov.cn/...
[SCOPSR-SHBB Script] 开始测试选择器...
[SCOPSR-SHBB Script] 当前网站配置: scopsr.gov.cn
```

### 3. 搜索页面测试步骤

#### SCOPSR 网站测试
1. 访问：https://www.scopsr.gov.cn/was5/web/search
2. 输入任意搜索词（如"管理"）并搜索
3. 检查是否出现右侧绿色的"批量保存"按钮
4. 在控制台查看调试信息：
   ```
   [SCOPSR-SHBB Script] 当前是搜索页面，测试搜索页面选择器...
   [SCOPSR-SHBB Script] 搜索容器元素: 找到
   [SCOPSR-SHBB Script] 文章项目数量: X
   ```

#### SHBB 网站测试  
1. 访问：https://www.shbb.gov.cn/search.jspx
2. 输入任意搜索词并搜索
3. 检查是否出现右侧绿色的"批量保存"按钮
4. 查看控制台调试信息

### 4. 文章详情页测试步骤
1. 从搜索结果中点击任一文章标题进入详情页
2. 检查是否出现右侧绿色的"保存为DOCX文档"按钮
3. 在控制台查看调试信息：
   ```
   [SCOPSR-SHBB Script] 当前是文章页面，测试文章页面选择器...
   [SCOPSR-SHBB Script] 文章内容元素: 找到，长度: XXXX
   [SCOPSR-SHBB Script] 文章标题元素: XXXXX
   ```

### 5. 功能测试
1. 点击"批量保存"或"保存为DOCX文档"按钮
2. 观察按钮状态变化和控制台日志
3. 正常情况下应该能看到详细的处理过程：
   ```
   [SCOPSR-SHBB Script] 开始收集当前页面文章链接...
   [SCOPSR-SHBB Script] 找到搜索结果容器
   [SCOPSR-SHBB Script] 找到 X 个文章元素
   [SCOPSR-SHBB Script] 开始获取文章内容: https://...
   [SCOPSR-SHBB Script] 文章请求成功，状态码: 200
   ```

## 常见问题诊断

### 问题1：脚本没有加载
**症状**：没有看到绿色的加载提示，没有保存按钮
**解决**：
1. 检查 Tampermonkey 是否启用
2. 检查脚本是否启用
3. 刷新页面重试

### 问题2：找不到搜索容器
**症状**：控制台显示"未找到搜索结果容器"
**诊断**：
1. 查看控制台中的"页面HTML结构预览"信息
2. 检查网站是否更新了页面结构
3. 网页可能还在加载中，等待几秒后再试

### 问题3：文章内容获取失败
**症状**：显示"HTTP错误"、"网络请求失败"或"请求超时"
**解决**：
1. 检查网络连接
2. 某些文章可能有访问限制
3. 降低保存速度（脚本已设置800ms延迟）

### 问题4：按钮显示但点击无反应
**症状**：按钮存在但点击后没有任何反应
**诊断**：
1. 查看控制台是否有JavaScript错误
2. 检查是否有其他脚本冲突
3. 尝试刷新页面重新测试

## 调试信息说明
- `[SCOPSR-SHBB Script]` 开头的都是脚本的调试信息
- 如果看到网络相关错误（如 `conac.cn` 或 `bdimg.share.baidu.com`），这些通常是网站自身的追踪脚本问题，不影响我们脚本的功能
- 重点关注脚本自身的日志信息

## 测试建议
1. 先在 SCOPSR 网站测试，因为它的结构相对稳定
2. 测试时选择较少页数（1-2页），避免过度请求
3. 如遇问题，完整复制控制台信息以便分析

## 支持的网站URL模式
- SCOPSR 搜索页：`https://www.scopsr.gov.cn/was5/web/search*`
- SCOPSR 文章页：`https://www.scopsr.gov.cn/xwzx/*`
- SHBB 搜索页：`https://www.shbb.gov.cn/search*`
- SHBB 文章页：`https://www.shbb.gov.cn/**/[数字].jhtml`