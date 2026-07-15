# CNIPA 专利信息批量查询

一个油猴（Tampermonkey）脚本，用于在[中国专利公布公告查询系统](https://cpquery.cponline.cnipa.gov.cn/chinesepatent/index)批量查询专利信息，并将结果导出为 Excel。

## 功能

- **6 种查询字段**（可勾选、可拖拽排序）：
  - 申请人
  - 代理机构
  - 最近缴费人
  - 最近缴费种类
  - 法律状态
  - 案件状态
- **两种输入方式**：直接粘贴申请号，或上传 xlsx 文件（内置模板下载）
- **导出 Excel**：查询结果一键导出 xlsx，只导出勾选的字段
- **批量控制**：暂停/继续、失败重查
- **申请号格式自适应**：支持 `CN302616662164.3`、`302616662164.3`、`3026166621643` 等格式，自动清洗为标准 13 位

## 安装

### 1. 安装 Tampermonkey 扩展

- Chrome / Edge：[Chrome 应用商店](https://chromewebstore.google.com/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)
- Firefox：[Firefox 附加组件](https://addons.mozilla.org/zh-CN/firefox/addon/tampermonkey/)

### 2. 安装本脚本

点击下面的链接，Tampermonkey 会弹出安装确认页，点"安装"即可：

**[安装 CNIPA 专利信息批量查询](https://raw.githubusercontent.com/SSI00/cnipa-userscript/main/cnipa_batch_query.user.js)**

## 使用方法

1. 打开 [CNIPA 专利查询页面](https://cpquery.cponline.cnipa.gov.cn/chinesepatent/index)并**登录你的账号**
2. 页面右上角会出现蓝色面板"CNIPA 专利信息批量查询"
3. **先在页面上手动搜索一次**（随便搜一个申请号），面板顶部变成"✅ 登录态已获取"后才能开始批量查询
4. 勾选要查询的字段（可拖动调整顺序）
5. 粘贴申请号（每行一个）或上传 xlsx 文件
6. 点"开始查询"，等待完成
7. 点"导出 Excel"下载结果

## 注意事项

- **必须先手动搜索一次**：脚本需要捕获页面请求中的登录令牌（JWT），不搜索一次无法调 API
- **个人账号只能查已公开的专利**：未公开的申请会显示"未公开"；代理机构账号可查本机构的未公开案件
- **查询期间不要关闭页面**：脚本依赖当前页面的登录态
- **登录态会过期**：如果提示"登录态过期"，在页面上重新搜索一次，然后点"重查失败项"
- **速度**：约 0.5~1 秒/条，几百条大概几分钟

## 更新

脚本已配置自动更新，Tampermonkey 会定期检查新版本并提示。也可以点面板标题栏的 ⓘ 图标查看更新日志。

## 常见问题

**Q: 面板没出现？**
A: 确认 Tampermonkey 已启用，且脚本在启用状态。刷新页面试试。

**Q: 一直提示"等待获取登录态"？**
A: 在页面的搜索框里随便搜一个申请号（甚至 `1111111111111`），等搜索结果出来后面板就会变绿。

**Q: 某些申请号查不到数据？**
A: 个人账号查不到未公开的申请，这是 CNIPA 的权限限制，不是脚本问题。

**Q: 导出的 Excel 里某些字段是空的？**
A: 该专利本身没有这项信息（如无代理机构），或该字段未公开。

## 许可

MIT
