# DOCX 条款交叉引用审查 API

基于 FastAPI + lxml 的 HTTP 服务：接收未加密的 DOCX，核对正文与表格中
**带书签的 Heading 1~3 自动编号标题**以及指向它们的 **REF 编号域**
（`\n`、`\r`、`\w` 开关），并针对文字与段落标记修订分别生成
**接受全部修订**与**拒绝全部修订**两个视图，报告每个视图中可见的每处 REF 域的
位置、现值、应值与状态。

## 快速启动

```bash
docker compose up --build
```

服务监听 `http://localhost:8000`：

- `POST /audit` — 审查接口
- `GET /` — 服务信息
- `GET /docs` — Swagger UI（可在线试调）

## 接口说明

### POST /audit

两种提交方式任选其一：

1. `multipart/form-data`，表单字段名 `file`；
2. 直接以请求体提交 DOCX 字节（Content-Type 任意）。

**curl 示例（multipart）：**

```bash
curl -s -X POST http://localhost:8000/audit \
  -F "file=@sample.docx"
```

**curl 示例（裸字节）：**

```bash
curl -s -X POST http://localhost:8000/audit \
  -H "Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document" \
  --data-binary @sample.docx
```

`sample.docx` 可由 `python3 examples/make_sample.py` 生成（仓库已附带
`examples/sample.docx` 的生成脚本）。

### 成功响应（200）

返回 `accept`（接受全部修订）与 `reject`（拒绝全部修订）两个视图，
各自列出该视图中可见的全部 REF 域。对 `examples/make_sample.py` 生成的
`sample.docx`（含一处被删除的标题、一处过期引用），实际响应为：

```json
{
  "views": {
    "accept": {
      "field_count": 3,
      "fields": [
        {
          "paragraph_index": 4,
          "field_index": 0,
          "paragraph_text": "见1.1（正确）",
          "instruction": "REF _TocScope \\w",
          "bookmark": "_TocScope",
          "switches": ["w"],
          "mode": "w",
          "current": "1.1",
          "expected": "1.1",
          "status": "correct"
        },
        {
          "paragraph_index": 5,
          "field_index": 0,
          "paragraph_text": "见1.3（过期：接受视图中应为 1.2）",
          "instruction": "REF _TocTerm \\w",
          "bookmark": "_TocTerm",
          "switches": ["w"],
          "mode": "w",
          "current": "1.3",
          "expected": "1.2",
          "status": "stale"
        },
        {
          "paragraph_index": 6,
          "field_index": 0,
          "paragraph_text": "见1.2（目标已删除）",
          "instruction": "REF _TocOld \\w",
          "bookmark": "_TocOld",
          "switches": ["w"],
          "mode": "w",
          "current": "1.2",
          "expected": null,
          "status": "target_missing"
        }
      ]
    },
    "reject": {
      "field_count": 3,
      "fields": [
        {
          "paragraph_index": 5,
          "field_index": 0,
          "paragraph_text": "见1.1（正确）",
          "instruction": "REF _TocScope \\w",
          "bookmark": "_TocScope",
          "switches": ["w"],
          "mode": "w",
          "current": "1.1",
          "expected": "1.1",
          "status": "correct"
        },
        {
          "paragraph_index": 6,
          "field_index": 0,
          "paragraph_text": "见1.3（过期：接受视图中应为 1.2）",
          "instruction": "REF _TocTerm \\w",
          "bookmark": "_TocTerm",
          "switches": ["w"],
          "mode": "w",
          "current": "1.3",
          "expected": "1.3",
          "status": "correct"
        },
        {
          "paragraph_index": 7,
          "field_index": 0,
          "paragraph_text": "见1.2（目标已删除）",
          "instruction": "REF _TocOld \\w",
          "bookmark": "_TocOld",
          "switches": ["w"],
          "mode": "w",
          "current": "1.2",
          "expected": "1.2",
          "status": "correct"
        }
      ]
    }
  }
}
```

#### 字段含义

| 字段 | 含义 |
| --- | --- |
| `paragraph_index` | 域所在段落在该视图可见段落序列中的序号（0 起，含表格内段落，按文档顺序） |
| `field_index` | 域在该段落内顶层域中的序号（0 起） |
| `paragraph_text` | 所在段落的可见文本预览（便于定位，最长 80 字符） |
| `instruction` | 域指令原文 |
| `bookmark` | 引用的书签名 |
| `switches` | 指令中的开关字母列表 |
| `mode` | 生效的取值开关（`n`/`r`/`w`，多个时后者优先；无开关时为 `null`，表示取书签文本） |
| `current` | 域的现值（域结果区文本；无结果区时为空串） |
| `expected` | 按 Word 规则重算的应值；目标缺失时为 `null` |
| `status` | `correct`（正确）/ `stale`（过期）/ `target_missing`（目标缺失） |

### 错误响应（422）

凡下列情况，整体返回 422，**不返回任何部分报告**：

```json
{
  "detail": {
    "error": "duplicate_bookmark",
    "message": "书签名重复：'_TocA'"
  }
}
```

| error | 触发条件 |
| --- | --- |
| `invalid_docx` | 请求体为空、无法解包（加密/损坏）、缺少 `word/document.xml` 或 `w:body` |
| `duplicate_bookmark` | 同一书签名出现两次及以上 |
| `broken_field` | 域结构损坏：开始/分隔/结束符不配对、指令文字在域外等 |
| `missing_r_context` | `\r` 引用之前（同一视图中）没有可见的编号标题作为相对上下文 |
| `unsupported_revision` | 含移动、属性变更、表格结构等修订类型（仅支持文字与段落标记的插入/删除） |

## 审查规则要点

- **编号来源**：段落直接 `numPr` → 样式继承链（`basedOn`）上的 `numPr` →
  编号级别 `pStyle` 链接；`numId=0` 视为显式取消编号。
- **列表实例**：计数器按 `numId` 独立维护；支持 `lvlOverride/startOverride`
  起始覆盖与级别整体替换；`lvlRestart` 缺省时任一更高级别出现后重启，
  `0` 表示永不重启。
- **标题识别**：样式名 `heading N` 或样式 ID `HeadingN`（含基于它们的
  自定义样式），仅 1~3 级参与编号上下文。
- **`\n`**：取目标标题当前级别的编号，按该级别 `numFmt` 格式化
  （decimal/decimalZero/upperRoman/lowerRoman/upperLetter/lowerLetter/ordinal）。
- **`\w`**：取目标标题的完整条号（如 `4.5.2`）。
- **`\r`**：以**引用点之前最近的可见编号标题**为上下文，去掉公共前缀，
  如源标题 `4.3.1` 指向 `4.5.2` 应显示 `5.2`；目标编号是上下文编号的
  前缀时回退为完整编号。无可用上下文时返回 422。
- **无开关**：REF 取书签文本比对。
- **双视图**：`w:ins` 内容仅计入接受视图，`w:del` 内容仅计入拒绝视图；
  段落标记删除/插入按 Word 规则与下一段合并（合并后取幸存段落标记所属
  段落的属性）。两视图独立计算，被删除的目标不会从另一视图借用。
- **跨 run**：域指令、域结果、书签起止跨 run 拆分时按文档顺序事件流配对。
- **表格**：表格（含嵌套）内段落与正文按文档顺序统一编号与定位。

## 本地开发与测试

```bash
pip install -r requirements.txt
pip install pytest httpx
python3 -m pytest tests/ -q
uvicorn app.main:app --reload
```

## 项目结构

```
app/
  main.py       # FastAPI 入口，POST /audit
  audit.py      # 审查编排：双视图、编号、REF 核对
  views.py      # 接受/拒绝视图构建，域结构解析
  numbering.py  # numbering.xml 解析与计数引擎
  styles.py     # styles.xml 解析、Heading 识别
  oxml.py       # DOCX 解包、修订类型与书签预检
  errors.py     # 业务错误（→ 422）
tests/          # pytest 端到端测试
examples/       # sample.docx 生成脚本
```
