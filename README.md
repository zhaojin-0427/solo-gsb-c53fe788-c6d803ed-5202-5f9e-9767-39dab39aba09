# DOCX 条款交叉引用审查 API

用 Python + FastAPI + lxml 实现的 HTTP 服务：接收**未加密 DOCX**，核对正文与表格中
**带书签的 Heading 1～3 自动编号标题**，以及指向它们的 **REF 编号域**，并分别生成
**接受修订**与**拒绝修订**两个视图下的审查结果。

每个视图独立报告其中**可见的每处编号 REF 域**的：

- 位置（段落序号、域序号、XPath、表格路径、段落文本）；
- 现值（域结果缓存的文本）；
- 应值（按 Word 编号规则重新计算的编号）；
- 状态：`correct`（正确）/ `stale`（过期）/ `missing_target`（目标缺失）。

---

## 1. 一键启动（Docker Compose）

```bash
docker compose up --build -d
```

服务监听 `http://localhost:8000`。查看日志：

```bash
docker compose logs -f
```

停止：

```bash
docker compose down
```

### 本地直接运行（可选）

```bash
python3 -m pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

---

## 2. 访问

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| `GET` | `/health` | 健康检查，返回 `{"status": "ok"}` |
| `POST` | `/audit` | 上传一个 `.docx` 文件（multipart 字段名 `file`），返回两视图审查报告 |

交互式文档（Swagger UI）：<http://localhost:8000/docs>

---

## 3. 请求示例

仓库自带示例文档 `samples/sample.docx`（由 `scripts/make_sample.py` 生成，
含多级编号、表内标题、跨 run 域、文字插入/删除修订与段落标记删除修订）。
若要重新生成：

```bash
python3 scripts/make_sample.py
```

### curl

```bash
curl -X POST http://localhost:8000/audit \
  -F "file=@samples/sample.docx;type=application/vnd.openxmlformats-officedocument.wordprocessingml.document"
```

### Python

```python
import requests

with open("samples/sample.docx", "rb") as f:
    resp = requests.post(
        "http://localhost:8000/audit",
        files={"file": ("sample.docx", f,
               "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
    )
print(resp.status_code)
report = resp.json()
```

### 响应示例（以下为对 `samples/sample.docx` 的真实输出摘录）

```json
{
  "filename": "sample.docx",
  "accept_view": {
    "fields": [
      {
        "position": {
          "paragraph_index": 4,
          "field_index": 0,
          "logical_paragraph_index": 4,
          "xpath": "/document/body/p[5]",
          "table_path": null,
          "paragraph_text": "见后文 5.1（接受视图应值 5.2，缓存已过期）；完整条号 4.5.2；本级编号 2"
        },
        "bookmark": "bm_t2",
        "switch": "r",
        "current_value": "5.1",
        "expected_value": "5.2",
        "status": "stale"
      },
      {
        "position": {
          "paragraph_index": 13,
          "field_index": 0,
          "logical_paragraph_index": 13,
          "xpath": "/document/body/tbl/tr/tc[2]/p",
          "table_path": "document/body/tbl/tr/tc",
          "paragraph_text": "表内引用 4.6.1"
        },
        "bookmark": "bm_tab",
        "switch": "w",
        "current_value": "4.6.1",
        "expected_value": "4.6.1",
        "status": "correct"
      },
      {
        "position": {
          "paragraph_index": 17,
          "field_index": 0,
          "logical_paragraph_index": 17,
          "xpath": "/document/body/p[16]",
          "table_path": null,
          "paragraph_text": "对旧条款的引用 4.8"
        },
        "bookmark": "bm_old",
        "switch": "w",
        "current_value": "4.8",
        "expected_value": null,
        "status": "missing_target"
      },
      {
        "position": {
          "paragraph_index": 19,
          "field_index": 0,
          "logical_paragraph_index": 18,
          "xpath": "/document/body/p[18]",
          "table_path": null,
          "paragraph_text": "（接受视图中本句并入上方标题）引用 4.9.1"
        },
        "bookmark": "bm_pilcrow",
        "switch": "w",
        "current_value": "4.9.1",
        "expected_value": "4.7.1",
        "status": "stale"
      }
    ]
  },
  "reject_view": {
    "fields": [
      {
        "position": {
          "paragraph_index": 4,
          "field_index": 0,
          "logical_paragraph_index": 4,
          "xpath": "/document/body/p[5]",
          "table_path": null,
          "paragraph_text": "见后文 5.1（接受视图应值 5.2，缓存已过期）；完整条号 4.5.2；本级编号 2"
        },
        "bookmark": "bm_t2",
        "switch": "r",
        "current_value": "5.1",
        "expected_value": "5.2",
        "status": "stale"
      },
      {
        "position": {
          "paragraph_index": 17,
          "field_index": 0,
          "logical_paragraph_index": 17,
          "xpath": "/document/body/p[16]",
          "table_path": null,
          "paragraph_text": "对旧条款的引用 4.8"
        },
        "bookmark": "bm_old",
        "switch": "w",
        "current_value": "4.8",
        "expected_value": "4.7",
        "status": "stale"
      },
      {
        "position": {
          "paragraph_index": 19,
          "field_index": 0,
          "logical_paragraph_index": 19,
          "xpath": "/document/body/p[18]",
          "table_path": null,
          "paragraph_text": "（接受视图中本句并入上方标题）引用 4.9.1"
        },
        "bookmark": "bm_pilcrow",
        "switch": "w",
        "current_value": "4.9.1",
        "expected_value": "4.7.1",
        "status": "stale"
      }
    ]
  }
}
```

字段含义：

| 字段 | 含义 |
| --- | --- |
| `position.paragraph_index` | 段落/表格单元格按文档顺序的 0 基物理序号（正文与表格统一编号） |
| `position.field_index` | 域在其所在段落内的 0 基序号 |
| `position.logical_paragraph_index` | 该域在当前视图合并段落标记后的逻辑段落序号 |
| `position.xpath` | 去除命名空间、带位置谓词的稳定 XPath |
| `position.table_path` | 位于表格内时为表格结构路径，否则为 `null` |
| `position.paragraph_text` | 所在段落纯文本片段（最多 120 字） |
| `bookmark` | REF 域引用的书签名 |
| `switch` | 编号开关：`n` / `r` / `w` |
| `current_value` | 域结果中的现值（跨 run 拼接） |
| `expected_value` | 重算的应值；目标在本视图缺失时为 `null` |
| `status` | `correct` / `stale` / `missing_target` |

### 错误响应（HTTP 422，不返回任何部分报告）

```json
{
  "error": {
    "code": "duplicate_bookmark",
    "message": "书签 \"bm_x\" 重复定义（段落 2 与 段落 5）"
  }
}
```

```bash
# 例如上传重复书签的文档
curl -i -X POST http://localhost:8000/audit -F "file=@bad.docx"
# HTTP/1.1 422 Unprocessable Entity
```

---

## 4. 编号与引用规则

### REF 开关

- `\n`：按**目标标题自身级别**的 `numFmt` 取本级编号。目标 4.5.2 → `2`。
- `\w`：按目标级别 `lvlText`（如 `%1.%2.%3`）取**完整条号**。目标 4.5.2 → `4.5.2`。
- `\r`：以**引用位置之前最近的可见自动编号标题**为上下文，取去掉共同前缀后的
  相对编号。上下文 4.3.1、目标 4.5.2 → 共同前缀 4，结果 **`5.2`**。
  - 上下文与目标属于不同列表实例（`w:numId` 不同）或无共同前缀时，等同完整条号；
  - 上下文与目标编号完全相同时，返回完整条号；
  - 若该视图中引用前不存在任何可见的 Heading 1～3 自动编号标题 → `422 no_r_context`。
- 允许随附良性开关 `\h` 与 `\* MERGEFORMAT`；其他开关（`\p \f \d \t` 及未知开关）→ `422`。
- 无编号开关的普通 REF 域不属于审查范围，不出现在报告中。

### 列表 / 样式 / 重启 / 起始覆盖（依 Word 规则）

- **列表实例**：每个 `w:num`（`w:numId`）维护独立计数器；同一 `abstractNum`
  被多个实例引用时计数互不影响。
- **样式继承**：标题级别沿 `w:basedOn` 链解析（内置 `Heading1/2/3`、`heading 1`、
  `标题 1` 等命名均可识别）；编号可在段落 `w:pPr/w:numPr` 直接声明，也可由样式链
  提供；直接 `w:numId w:val="0"` 取消样式编号。
- **级别重启**：更高级别编号继续递增时，按各级 `w:lvlRestart` 重置更深级别
  （`0`=任何更高级别均重启；`k`=仅最近 `k` 级）。级别首次出现或被重置后重新开始，
  不会连带重启其子级（与 Word 一致）。
- **起始覆盖**：`w:start` 提供级别初始值；`w:lvlOverride/w:startOverride`
  仅对该实例该级别的**第一次出现**生效，之后按级别 `start`；接受/拒绝视图各跑一遍
  计数，覆盖各自独立消费。
- **格式**：支持 `decimal`、`decimalZero`、`upperRoman/lowerRoman`、
  `upperLetter/lowerLetter`、`ordinal`、`none`；`w:isLgl` 强制更高级别按十进制渲染；
  遇到 `bullet`、中文等暂不支持的 `numFmt` 返回 `422 unsupported_number_format`。

### 修订视图

- **文字修订**：`w:ins` 包裹的 run/域/书签仅存在于接受视图；`w:del` 包裹的
  （含 `w:delText`）仅存在于拒绝视图。两视图独立计数、独立判定目标，
  **删除目标不会从另一视图借用**。
- **段落标记修订**：`w:pPr/w:rPr/w:del` 接受后段落并入后一段（宿主取后段属性）；
  `w:pPr/w:rPr/w:ins` 拒绝后段落并入前一段（宿主取前段属性）；位置归属按
  Word 合并方向处理。
- **跨 run 对应**：复杂域（`w:fldChar` begin/separate/end + `w:instrText`）与
  `w:fldSimple` 均支持；指令与结果可分散在多个 run 中，按出现顺序拼接。
  书签 `bookmarkStart`/`bookmarkEnd` 同理。一个域的全部 run 必须处于同一修订容器
  （或都不在容器中），否则视为结构损坏。

### 422 拒绝条件（拒绝整个请求，无部分报告）

| code | 触发条件 |
| --- | --- |
| `not_a_docx` | 非 OOXML ZIP、旧版 `.doc`、缺少关键部件 |
| `encrypted_docx` | 声明为加密包或含 EncryptedPackage |
| `corrupt_docx` | 关键 XML 无法解析、缺少主文档 |
| `duplicate_bookmark` | 同名 `bookmarkStart` 重复 |
| `malformed_field` | 域 begin/separate/end 不配对、跨修订容器、指令区混入普通文本、REF 缺书签或多编号开关 |
| `unsupported_ref_switch` | REF 含不支持的开关 |
| `unsupported_revision` | 除文字插入/删除、段落标记插入/删除（容忍 `w:pPrChange`/`w:rPrChange`）外的任何修订：移动、表格/单元格/行修订、编号变更、批注等 |
| `conflicting_pilcrow_revision` | 同一段落标记同时存在插入与删除修订 |
| `no_r_context` | 某视图中 `\r` 域之前没有任何可见编号标题 |
| `unsupported_number_format` / `unsupported_numbering_level` / `unknown_numbering_instance` 等 | 编号定义无法按规则计算 |

---

## 5. 项目结构

```
app/
  main.py                     FastAPI 入口（POST /audit、GET /health）
  docx_audit/
    package.py                DOCX 解包、加密检测、关系部件解析
    namespaces.py / errors.py 命名空间与 422 错误模型
    revisions.py              修订类型白名单校验
    tokens.py                 段落分词：run / 复杂域 / fldSimple / 书签（跨 run）
    fields.py                 REF 指令与开关解析
    styles.py                 样式继承、Heading 判定
    numbering.py              列表实例、级别重启、起始覆盖、n/r/w 渲染
    auditor.py                主流程：物理段 → 接受/拒绝视图 → 编号 → 域求值
tests/                        pytest 套件（编号引擎、视图、HTTP、422 场景）
scripts/make_sample.py        示例文档生成器
samples/sample.docx           示例 DOCX
Dockerfile / docker-compose.yml
```

## 6. 运行测试

```bash
python3 -m pip install -r requirements.txt pytest httpx
python3 -m pytest -q
```
