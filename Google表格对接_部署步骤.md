# Google 表格对接部署步骤（Apps Script Web App）

> 作用：让 `KOL_Scout_v3.html` 每次搜索都能和长期建联表「达人收集池」去重，并在你点「加入」时自动把新达人写回这张表。换电脑、清缓存都不会丢去重记录。
>
> 这份代码需要**你本人**在自己的 Google 账号里粘贴、部署（Claude 无法替你操作 Google 账号）。照下面 7 步做即可，全程约 5 分钟。

---

## 一、整体原理（一句话）

工具（前端）→ 调用你部署的 Web App（Apps Script）→ 读/写你的 Google 表格。
表格保持私有，Web App 以「你本人身份」运行，所以静态托管在 GitHub 上的工具也能读写私有表。

```
KOL_Scout_v3.html
   │  GET  →  取「达人收集池」账号链接列（用于去重）
   │  POST →  追加新达人（写回）
   ▼
Apps Script Web App（你部署，含口令校验）
   ▼
Google Sheet「达人收集池」
```

---

## 二、Apps Script 代码（整段复制）

> 部署前**只需改一处**：把 `TOKEN` 改成你自己的口令（任意字符串，别人猜不到即可），稍后要把同样的口令填进工具。
>
> ⚠️ **已经部署过旧版的**：这份代码新增了「批量转换历史链接」所需的两段逻辑（`doGet` 的 `mode=rows`、`doPost` 的 `action:'updateUrls'`）。请**整段替换**旧代码，然后**必须重新部署一次新版本**（`部署 → 管理部署 → 编辑(铅笔) → 版本选「新版本」→ 部署`），否则新功能不生效。**URL 不变**，工具里无需改动。

```javascript
// ===== KOL 建联表对接 Web App =====
// 部署前把下面的 TOKEN 改成你自己的口令，并把同样的口令填进工具的「口令 Token」框
const TOKEN      = 'change_me_到你自己的口令';
const SHEET_NAME = '达人收集池';   // 目标工作表名（sheet3）
const URL_COL    = 4;              // 「账号链接」在第 4 列

// 读取：默认返回「账号链接」列所有值（供去重）；?mode=rows 时带行号返回（供批量转换历史链接）
function doGet(e){
  try{
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if(!sh) return json({ok:false, error:'找不到工作表：'+SHEET_NAME});
    const last = sh.getLastRow();
    const mode = (e && e.parameter && e.parameter.mode) || '';

    if(mode === 'rows'){
      // 带行号：[{row, url}]，供前端把 @handle 转成 /channel/UCxxxx 后按行回写
      let rows = [];
      if(last >= 2){
        const vals = sh.getRange(2, URL_COL, last-1, 1).getValues();
        vals.forEach(function(r, i){
          rows.push({ row: i + 2, url: (r[0]||'').toString().trim() });
        });
      }
      return json({ok:true, rows:rows});
    }

    let urls = [];
    if(last >= 2){
      urls = sh.getRange(2, URL_COL, last-1, 1).getValues()
               .map(r => (r[0]||'').toString().trim())
               .filter(Boolean);
    }
    return json({ok:true, urls:urls});
  }catch(err){
    return json({ok:false, error:String(err)});
  }
}

// 写回：校验口令后，默认追加新达人；action='updateUrls' 时按行号更新账号链接列
function doPost(e){
  try{
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if(body.token !== TOKEN) return json({ok:false, error:'unauthorized'});

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if(!sh) return json({ok:false, error:'找不到工作表：'+SHEET_NAME});

    // 批量更新账号链接（把历史 @handle 链接转成 /channel/UCxxxx）
    if(body.action === 'updateUrls'){
      const updates = body.updates || [];
      let updated = 0;
      updates.forEach(function(u){
        if(u && u.row >= 2 && u.url){
          sh.getRange(u.row, URL_COL).setValue(u.url);
          updated++;
        }
      });
      return json({ok:true, updated:updated});
    }

    const rows = body.rows || [];
    let added = 0;
    rows.forEach(function(r){
      sh.appendRow([
        r.email    || '',        // 1  Contact Email
        '',                      // 2  编号
        r.name     || '',        // 3  ChannelName
        r.url      || '',        // 4  账号链接
        r.platform || 'YouTube', // 5  Platform
        r.country  || '',        // 6  国家/地区
        '',                      // 7  内容垂类/风格
        r.followers|| '',        // 8  数据指标 粉丝量
        r.avgViews || '',        // 9  平均播放量
        (r.er !== '' && r.er != null) ? r.er : '', // 10 互动率
        r.rating   || '',        // 11 综合评级
        '',                      // 12 代理人姓名
        '',                      // 13 代理人邮箱
        '',                      // 14 是否适合合作
        '',                      // 15 建议合作内容
        '',                      // 16 建议产品预算
        r.date     || '',        // 17 建联日期
        '未建联',                // 18 当前状态
        '',                      // 19 备注
        '', '', '', ''           // 20-23 FirstName / Status / SentDate / ReplyStatus
      ]);
      added++;
    });
    return json({ok:true, added:added});
  }catch(err){
    return json({ok:false, error:String(err)});
  }
}

function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}
```

---

## 三、部署 7 步（照做即可）

**第 1 步｜打开脚本编辑器**
打开你的表格 → 顶部菜单 `扩展程序` → `Apps Script`。会新开一个脚本项目页面。

**第 2 步｜粘贴代码**
把编辑器里默认的 `function myFunction(){}` 全部删掉 → 粘贴上面「二、Apps Script 代码」整段 → 把 `TOKEN` 改成你自己的口令 → 按 `Ctrl+S` 保存（可给项目起名，如「KOL建联表对接」）。

**第 3 步｜发起部署**
右上角点 `部署` → `新建部署`。

**第 4 步｜选类型**
点齿轮图标 `选择类型` → 选 `Web 应用 / Web app`。

**第 5 步｜配置访问权限**
- `说明`：随便填，如 v1
- `执行身份 / Execute as`：选 **我（你自己）**
- `谁可以访问 / Who has access`：选 **任何人 / Anyone**
  （必须选「任何人」，否则工具无法免登录调用。安全靠代码里的 TOKEN 兜底）
点 `部署`。

**第 6 步｜授权**
首次部署会弹授权 → 点 `授权访问` → 选你的 Google 账号 → 若提示「Google 未验证此应用」，点 `高级` → `转至〈项目名〉(不安全)` → `允许`。（这是给你自己的脚本授权，正常现象）

**第 7 步｜复制 Web App URL**
部署成功后会显示一个网址，形如：
```
https://script.google.com/macros/s/AKfyc..................../exec
```
点 `复制`。**这个 `/exec` 结尾的网址就是要填进工具的 URL。**

---

## 四、填进工具

打开 `KOL_Scout_v3.html` → 【达人搜集】Tab → 拉到底部「建联表（Google Sheet）对接」：
1. `Web App URL`：粘贴上一步复制的 `/exec` 网址
2. `口令 Token`：填你在代码里设的 `TOKEN`（要完全一致）
3. 点 `测试连接` → 显示「✓ 连接成功，建联表已有 N 条记录」即成功

之后去【批量搜索】Tab 搜索，结果里已在表内的会标 `📇 已在建联表`；点「加入」会自动写回表格。

---

## 五、批量转换历史链接（把旧 @handle 转成 /channel/UCxxxx）

**为什么要转**：去重靠 channelId（`UCxxxx`）比对最准。表里老数据若是 `youtube.com/@name` 形式，反推不出 channelId，可能造成重复达人。这个功能一次性把它们批量转成 `/channel/UCxxxx`。

**前置条件**：
1. 已按本文档部署**新版** Apps Script 并**重新部署新版本**（见第二节的红字提示）。
2. 工具里已填好 `Web App URL`、`口令 Token`、`API Key`。
3. **建议先备份表格**（`文件 → 制作副本`），以防万一。

**操作**：
【达人搜集】Tab → 底部「建联表对接」→ 点 `扫描并转换历史链接`。它会：
- 读取表里所有账号链接（带行号）；
- 只处理 `@handle`、`/user/xxx` 两种能反查的形式，逐条用 YouTube API 解析出 channelId；
- 按行号把第 4 列改写成 `https://www.youtube.com/channel/UCxxxx`；
- 结束显示「成功 N 条 / 无法转换 M 条（清单）」。

**注意**：
- 每转 1 条约耗 **1 次** YouTube API 配额；配额用尽会自动停下并写回已解析部分，改天可再点一次继续（已是 `/channel/` 的会自动跳过）。
- `/c/xxx` 自定义链接、TikTok/IG 链接、已改名/删号的频道**无法自动转换**，会列进清单，需手动处理或忽略。
- 只改动第 4 列「账号链接」，不动其它任何列。

---

## 六、常见问题

| 现象 | 原因 / 解决 |
|---|---|
| 测试连接报 `找不到工作表：达人收集池` | 脚本 `SHEET_NAME` 和表格里的工作表名不一致，改成实际名字 |
| 写回报 `unauthorized` | 工具里的口令和代码里的 `TOKEN` 不一致 |
| 测试连接失败 / 网络错误 | URL 没复制全（必须 `/exec` 结尾）；或第 5 步没选「任何人」 |
| 改了代码不生效 | 每次改代码要重新 `部署 → 管理部署 → 编辑(铅笔) → 版本选「新版本」→ 部署`，URL 不变 |
| 去重不够准 | 「账号链接」列若是 `/@handle`、`/c/xxx` 或 TikTok/IG 链接，无法反推 channelId，只能按链接原文比对，可能漏判；`/channel/UCxxxx` 形式最准。可用「批量转换历史链接」把 `@handle`/`user` 批量转成 `/channel/` |
| 点「扫描并转换历史链接」报错/无反应 | 多半是没重新部署新版代码；按第二节红字重新部署新版本。或没填 API Key（解析 channelId 需要） |

---

## 七、字段映射对照（工具 → 表格 23 列）

| 表格列 | 写入内容 |
|---|---|
| Contact Email | 搜到的邮箱（可能为空） |
| ChannelName | 频道名 |
| 账号链接 | 频道 URL |
| Platform | YouTube |
| 国家/地区 | 频道注册国 |
| 数据指标 粉丝量 | 订阅数 |
| 平均播放量 | 近 10 条均播放 |
| 互动率 | 计算值（%） |
| 综合评级 | ⭐ 评级 |
| 建联日期 | 加入当天日期 |
| 当前状态 | 未建联 |
| 其余列（编号/垂类/代理人/备注/FirstName 等） | 留空，供你后续人工补 |
