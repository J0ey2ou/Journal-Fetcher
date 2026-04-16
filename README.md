# Journal Fetcher

推荐使用新文件：

`journal_fetcher_allinone.py`

## 这一版新增了什么

- 主题可留空
- 留空主题时抓取该时间段内该期刊的所有文章
- 单文件版本
- 默认导出 Excel 到脚本同目录
- 本地网页可视化入口
- 页面内直接显示说明
- 只需要 Python 3，无第三方依赖

## 终端模式

```powershell
python journal_fetcher_allinone.py
```

## 网页模式

```powershell
python journal_fetcher_allinone.py --web
```

默认地址：

`http://127.0.0.1:8765`
