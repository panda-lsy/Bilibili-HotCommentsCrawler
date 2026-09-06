# Bilibili 评论快照准备工具

这个仓库最初是一个 B 站网页接口评论爬虫。为了用于 LocalPulse Edge 比赛版，当前仓库已把主路径调整为：

> **用户自行取得并导出的本地评论快照 → 脱敏与字段规范化 → LocalPulse 离线分析**

它不负责替用户申请平台权限，也不把实时爬虫作为比赛 Demo 依赖。

## 比赛模式

比赛模式只读取本地文件：

```powershell
python prepare_competition_data.py `
  --input result/comment_output.xlsx `
  --output localpulse_comments.csv `
  --platform Bilibili
```

支持：

- CSV：UTF-8、UTF-8 with BOM、GB18030
- JSON：数组，或包含 `comments` / `data` / `records` / `items` 数组的对象
- XLSX：逐工作表读取，适配仓库历史导出的多视频工作簿

输出 CSV 只保留 LocalPulse 所需字段：

```text
id, platform, content, likes, timestamp
```

转换器会做以下处理：

- 清理空评论、规范空白字符、去除转换后完全重复的记录；
- 解析点赞数和时间；
- 脱敏评论中的手机号、邮箱和 URL；
- 丢弃昵称、性别、IP、UID、用户 ID、回复对象和用户等级；
- 旁边生成 `<输出文件>.provenance.json`，记录来源文件哈希、行数和处理摘要。

注意：转换器的脱敏和本地处理不等于自动取得数据授权。比赛使用前，仍须确认数据来源、使用范围和主办方规则。

## 与 LocalPulse Edge 配合

将生成的 `localpulse_comments.csv` 交给 LocalPulse Edge：

```powershell
cd path\to\localpulse-edge
python scripts/run_demo.py path\to\localpulse_comments.csv
```

生成的 provenance 文件不要和原始导出文件一起提交；原始文件也不应提交到公开仓库。

## 旧实时脚本

`bili_comment.py` 保留历史实现，便于追溯，但已默认禁用。它涉及网页内部接口、批量获取和用户字段，不属于比赛主路径。

只有在已经取得明确书面授权的研究场景中，才可以显式设置：

```powershell
$env:BILI_ALLOW_LEGACY_LIVE = '1'
python bili_comment.py
```

B 站开放平台服务协议对未经书面同意使用机器人、蜘蛛或爬虫程序获取平台数据有明确限制，请以当前官方协议和实际授权为准：

<https://open.bilibili.com/agreement/developer-service>

## 许可证

本仓库保留原有许可证与历史文件。上游数据的版权、个人信息和使用授权不因代码开源而自动获得。
