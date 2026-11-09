# Contributing

请用小范围修改扩展规则，并在 pull request 中说明测试使用的输入类型、操作系统、生成页数和耗时。

## 修改位置

- 文件发现、四首歌识别、报告图片配对：`skill/scripts/input_manifest.py`
- 经文获取及缓存：`skill/scripts/spring_bible.py`
- 排版、页序、macOS/Windows PowerPoint 后端：`skill/scripts/combine_sunday_ppt.py`
- 成品对照规则：`skill/scripts/validate_against_reference.py`
- AI 执行流程：`skill/SKILL.md`

## 必须保持的约束

- 敬拜、讲道和回应诗歌必须保留为原生 PowerPoint 页面。
- 经文必须是可编辑文字。
- 二维码与对应报告同页，海报和 PDF 单独成页。
- 三一颂背景必须保留，停车通知页必须隐藏。
- 每次生成必须执行原生页面内部验证。

## 测试

```bash
python -m unittest discover -s tests
python skill/scripts/combine_sunday_ppt.py /sample/folder --dry-run
```

macOS 和 Windows 后端生成相同的逻辑页序，但必须分别在装有 Microsoft PowerPoint 的电脑上测试原生复制。

