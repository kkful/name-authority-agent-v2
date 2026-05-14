"""V2 第一步：论文采集器 — 搜作者 → 打开HTML → 写入Excel

用法:
    python collector.py 张伟
    python collector.py 张伟 --max 10  # 最多打开10篇
"""
import sys, os, json, time, re
import urllib.request
import pandas as pd

# 复用V1的CDP客户端
sys.path.insert(0, r"E:\名称规范系统\旧规范文档")
from author_agent.cdp_client import new_tab, close_tab, eval_js, click_at, page_text

PROXY = "http://localhost:3456"
OUTPUT_DIR = r"E:\名称规范系统\新规范文档系统\collected"

def get(path):
    with urllib.request.urlopen(PROXY + path) as r:
        return json.loads(r.read())

def post(path, data):
    body = data.encode("utf-8")
    req = urllib.request.Request(PROXY + path, data=body, method="POST")
    req.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(req, timeout=120) as r:
        return json.loads(r.read())

def find_or_create_cnki_tab():
    """找已激活的知网标签,没有则创建"""
    tabs = get("/targets")
    # 先找已激活的(有Classid)
    for t in tabs:
        if "AdvSearch" in t.get("url", ""):
            try:
                r = eval_js(t["targetId"], "JSON.parse(window.cnkiSearch.getSearchJsonInfo()).Classid")
                if r and str(r).strip():
                    return t["targetId"]
            except: pass
    # 没有则创建新标签,提示激活
    for t in tabs:
        if "AdvSearch" in t.get("url", ""):
            try: get("/close?target=" + t["targetId"])
            except: pass
    r = get("/new?url=https://kns.cnki.net/kns8s/AdvSearch")
    time.sleep(4)
    post(f"/eval?target={r['targetId']}",
        'var d=document.createElement("div");d.style.cssText="position:fixed;top:20px;left:50%;transform:translateX(-50%);background:#27ae60;color:#fff;padding:20px 40px;font-size:18px;z-index:999999;border-radius:8px;text-align:center";d.innerHTML="<b>请手动搜索一次激活页面</b><br>切换到「作者发文检索」→ 填入任意搜索条件 → 点检索";document.body.appendChild(d);"ok"')
    print("   -> 请在Chrome知网页面手动搜索一次激活,然后重新运行采集器")
    sys.exit(0)

def search_author(tab_id, author_name, max_results=200):
    """同步XHR搜索作者, 正确翻页（完整QueryJson + turnpage令牌）

    关键: QueryJson必须包含Resource/Products/KuaKuCode等完整字段;
          翻页需要turnpage令牌 + 18个完整POST参数;
          每页的详情链接按论文顺序累积，供采集阶段使用
    """
    all_titles = []
    all_detail_urls = []  # 每页的详情链接，按论文顺序累积
    total_count = 0
    brief_request = {}
    turnpage_token = ""

    for page in range(1, 20):
        if page == 1:
            # 第1页: 构建完整QueryJson
            js = f'''
var cs=window.cnkiSearch;
var base=JSON.parse(cs.getSearchJsonInfo());
base.Resource="CROSSDB";
base.Classid="WD0FTY92";
base.ExScope="1,2,3,4,5,6,7,8";
base.Rlang="CHINESE";
base.SearchType=3;
base.Products="CJFQ,CAPJ,ZHYX,CJTL,CDFD,CMFD,CPFD,IPFD,CPVD,CCND,WBFD,SCSF,SCHF,SCSD,SNAD,CCJD,CJFN,CCVD";
base.KuaKuCode="YSTT4HG0,LSTPFY1C,JUP3MUPD,MPMFIG1A,EMRPGLPA,WQ0UVIAA,BLZOG7CK,PWFIRAGL,NN3FJMUV,NLBO1Z6R";
base.QNode.QGroup=[{{Key:"S",Title:"",Logic:0,Items:[],ChildItems:[]}}];
base.QNode.QGroup[0].ChildItems=[{{Key:"a",Title:"",Logic:0,Items:[{{Key:"a",Title:"",Logic:0,Field:"AU",Operator:"DEFAULT",Value:"{author_name}",Value2:""}}],ChildItems:[]}}];
cs.setQueryJson(base);
var qj=JSON.stringify(base);
var x=new XMLHttpRequest();
x.open("POST","/kns8s/brief/grid",false);
x.setRequestHeader("Content-Type","application/x-www-form-urlencoded;charset=UTF-8");
x.send("boolSearch=true&QueryJson="+encodeURIComponent(qj));
window.__h=x.responseText;
"ok"
'''
        else:
            if not brief_request or not turnpage_token:
                break
            inner_qj = json.loads(brief_request['queryJson'])
            inner_qj['pageNum'] = page
            inner_qj['pageSize'] = 20
            inner_qj_str = json.dumps(inner_qj)
            js = f'''
var params={{
    boolSearch:"false",
    QueryJson:{json.dumps(inner_qj_str)},
    pageNum:"{page}",
    pageSize:"20",
    sortField:{json.dumps(brief_request.get("sortField","PT"))},
    sortType:"desc",
    dstyle:"listmode",
    boolSortSearch:"false",
    sentenceSearch:"false",
    productStr:{json.dumps(brief_request.get("productStr",""))},
    aside:"",
    searchFrom:{json.dumps("资源范围：总库;  时间范围：发表时间：不限;  ")},
    manageId:"",
    subject:"",
    turnpage:{json.dumps(turnpage_token)},
    language:"",
    uniplatform:""
}};
var parts=[];
for(var k in params){{parts.push(k+"="+encodeURIComponent(params[k]));}}
var body=parts.join("&");
var x=new XMLHttpRequest();
x.open("POST","/kns8s/brief/grid",false);
x.setRequestHeader("Content-Type","application/x-www-form-urlencoded;charset=UTF-8");
x.send(body);
window.__h=x.responseText;
"ok"
'''

        eval_js(tab_id, js)
        html = str(eval_js(tab_id, 'window.__h'))
        if len(html) < 100: break
        titles = re.findall(r'<a[^>]*class="fz14"[^>]*>([^<]+)</a>', html)
        titles = [t.strip() for t in titles if len(t.strip())>3 and t.strip() not in ('题名','作者','来源','发表时间','数据库','被引','下载','操作')]

        if page == 1:
            m = re.search(r'共找到</span>\s*<em>([\d,]+)</em>', html)
            total_count = int(m.group(1).replace(',','')) if m else 0
            # 保存第1页HTML用于DOM注入
            eval_js(tab_id, 'window.__h_page1=window.__h')

        # 提取briefRequest和turnpage令牌
        br_match = re.search(r'id="briefRequest"[^>]+value="([^"]+)"', html)
        if br_match:
            br_val = br_match.group(1)
            br_val = br_val.replace('&quot;','"').replace('&lt;','<').replace('&gt;','>').replace('&amp;','&')
            try:
                brief_request = json.loads(br_val)
            except:
                brief_request = {}

        tp_match = re.search(r'id="hidTurnPage"[^>]+value="([^"]+)"', html)
        if tp_match:
            turnpage_token = tp_match.group(1)

        all_titles.extend(titles)
        # 提取本页的详情链接（按论文顺序）
        page_urls = re.findall(r'/kcms2/article/abstract\?v=[^"&\']+', html)
        if not page_urls:
            page_urls = re.findall(r'/kcms2/article/abstract\?v=[^"\']+', html)
        all_detail_urls.extend(page_urls)
        print(f"  第{page}页: {len(titles)}篇, {len(page_urls)}个详情链接, 累计{len(all_titles)}篇")
        if len(all_titles) >= max_results: break
        if len(titles) == 0: break
        if not brief_request: break

    # 恢复第1页HTML用于DOM注入
    eval_js(tab_id, 'if(window.__h_page1)window.__h=window.__h_page1')
    return {"count": total_count, "titles": all_titles[:max_results], "len": len(all_titles), "detail_urls": all_detail_urls[:max_results]}

def collect_all_papers(tab_id, author_name, titles, detail_urls=None, max_open=20):
    """全量采集: 有HTML则打开全文, 无HTML则从预提取的详情链接获取。每轮重新注入DOM"""
    # 从API返回HTML中解析全部论文元数据(不截断)
    meta_js = 'var html=window.__h;' + \
              'var papers=[];' + \
              'var rows=html.split("<tr");' + \
              'rows.forEach(function(r){' + \
              'var t=r.replace(/<[^>]+>/g," ").replace(/\\s+/g," ").trim();' + \
              'if(t.length>20&&t.indexOf("题名")===-1&&t.indexOf("操作")===-1)papers.push(t.substring(0,300));' + \
              '});' + \
              'JSON.stringify(papers)'
    r = eval_js(tab_id, meta_js)
    try: metadata = json.loads(str(r))
    except: metadata = []

    MAX_PAPERS = 200
    # 优先用翻页后的titles总数, metadata只是第一页
    actual_total = max(len(titles), len(metadata))
    total = min(actual_total, MAX_PAPERS)
    if actual_total > MAX_PAPERS:
        print(f"  搜索结果 {actual_total} 条, 取前{MAX_PAPERS}篇")
    else:
        print(f"  搜索结果 {total} 条, 全部采集")

    # 注入DOM
    eval_js(tab_id, 'var bb=document.getElementById("briefBox");if(bb){bb.innerHTML=window.__h};"ok"')
    time.sleep(0.5)
    r = eval_js(tab_id, 'var c=0;document.querySelectorAll("a").forEach(function(a){if(a.textContent.indexOf("HTML阅读")>-1)c++});String(c)')
    html_count = int(r) if r and str(r).isdigit() else 0
    print(f"  其中 {html_count} 篇有HTML阅读")

    papers = []
    html_opened = 0  # 已打开的HTML篇数
    for i in range(total):
        idx = i + 1
        try:
            paper = {"序号": idx, "论文元数据": metadata[i] if i < len(metadata) else "",
                     "论文标题": titles[i] if i < len(titles) else metadata[i][:50] if i < len(metadata) else "",
                     "作者简介原文": "", "HTML全文": "", "全文长度": 0, "来源": "仅元数据(无HTML阅读)"}

            # 每轮重新注入DOM（防止上轮点击后DOM被清）
            if i > 0:
                eval_js(tab_id, 'var bb=document.getElementById("briefBox");if(bb){bb.innerHTML=window.__h};"ok"')
                time.sleep(0.3)

            # 分支处理：有HTML开HTML, 无HTML开详情页
            is_html = False
            has_html_link = i < html_count
            can_open_html = has_html_link and html_opened < max_open

            if can_open_html:
                eval_js(tab_id, f'var c=0;document.querySelectorAll("a").forEach(function(a){{if(a.textContent.indexOf("HTML阅读")>-1){{c++;if(c=={i+1})a.id="__ht{i+1}"}}}});"m"')
                time.sleep(0.2)

                click_body = f"#__ht{i+1}".encode("utf-8")
                req = urllib.request.Request(f"{PROXY}/clickAt?target={tab_id}", data=click_body, method="POST")
                req.add_header("Content-Type", "text/plain")
                with urllib.request.urlopen(req, timeout=30) as rp:
                    cr = json.loads(rp.read())
                if cr.get("clicked"):
                    time.sleep(3)
                    tabs = get("/targets")
                    html_tab = ""
                    for t in tabs:
                        if "HTML" in t.get("title","") or "reader" in t.get("url",""):
                            html_tab = t["targetId"]
                    if html_tab:
                        full_text = page_text(html_tab)
                        if full_text and len(full_text) > 100:
                            # 精准提取：标题+作者机构+摘要+关键词+基金+作者简介
                            txt = full_text
                            result = {}
                            # 标题(第一行)
                            result["标题"] = txt.split("\n")[0].strip()[:100]
                            # 摘要:从"摘要"到"关键词"或"Abstract"（排除英文噪音）
                            abs_m = re.search(r'(?:中文\s*)?摘要[：:]\s*([\s\S]+?)(?:关键词|Key\s*words|Abstract|基金|收稿|$)', txt)
                            if not abs_m:
                                abs_m = re.search(r'摘要[：:]\s*([\s\S]+?)(?:关键词|Key\s*words|Abstract|基金|收稿|$)', txt)
                            if abs_m: result["摘要"] = abs_m.group(1).strip()[:800]
                            # 关键词
                            kw_m = re.search(r'(关键词[：:][^\n]+)', txt)
                            if kw_m: result["关键词"] = kw_m.group(1).strip()[:200]
                            # 基金
                            fund_m = re.search(r'(基金[：:][^\n]+)', txt)
                            if fund_m: result["基金"] = fund_m.group(1).strip()[:200]
                            # 作者简介
                            bio_m = re.search(r'作者简介[：:]\s*([\s\S]+?)(?:收稿日期|Received|$)', txt)
                            if bio_m:
                                result["作者简介"] = bio_m.group(1).strip()[:500]
                                paper["作者简介原文"] = bio_m.group(1).strip()[:500]
                            # 收稿日期
                            date_m = re.search(r'(收稿日期[：:][^\n]+)', txt)
                            if date_m: result["收稿日期"] = date_m.group(1).strip()[:50]
                            # 拼接
                            parts = []
                            for k,v in result.items():
                                if v: parts.append(f"{k}:{v}")
                            paper["页面头部信息"] = "\n---\n".join(parts)
                            paper["来源"] = "HTML阅读"
                            is_html = True
                            html_opened += 1
                        close_tab(html_tab)

            # 非HTML论文：从预提取的detail URL获取（所有页面的链接已汇总）
            if not is_html and not paper.get("页面头部信息"):
                if detail_urls and i < len(detail_urls):
                    detail_url = "https://kns.cnki.net" + detail_urls[i].replace("&amp;","&")
                    r = get("/new?url=" + urllib.request.quote(detail_url, safe=':/?=&'))
                    d_tab = r.get("targetId","")
                    if d_tab:
                        time.sleep(3)
                        d_txt = page_text(d_tab)
                        if d_txt and len(d_txt) > 100:
                            result = {}
                            # 跳过CNKI页头"总库"等,取真正的标题
                            lines = d_txt.split("\n")
                            title = ""
                            for line in lines:
                                s = line.strip()
                                if s and s not in ("总库","检索","CNKI AI","") and "http" not in s and len(s) > 5:
                                    title = s[:100]; break
                            result["标题"] = title
                            # 作者+机构(标题后到摘要前)
                            abs_idx = d_txt.find("摘要")
                            if abs_idx > -1:
                                author_block = d_txt[len(title):abs_idx].strip()[:500]
                                if author_block:
                                    result["作者机构"] = author_block.replace("\n"," ")[:400]
                            # 摘要
                            abs_m = re.search(r'摘要[：:]\s*([\s\S]+?)(?:关键词|Key\s*words|基金|专辑|专题|分类|$)', d_txt)
                            if abs_m: result["摘要"] = abs_m.group(1).strip()[:800]
                            # 关键词(去重复前缀)
                            kw_m = re.search(r'关键词[：:]\s*([^\n]+)', d_txt)
                            if kw_m: result["关键词"] = "关键词:" + kw_m.group(1).strip()[:200]
                            # 基金
                            fund_m = re.search(r'基金[资助]?[：:]\s*([^\n]+)', d_txt)
                            if fund_m: result["基金"] = "基金:" + fund_m.group(1).strip()[:200]
                            # 专辑/专题/分类号
                            for field in ["专辑","专题","分类号","在线公开时间"]:
                                fm = re.search(rf'{field}[：:]\s*([^\n]+)', d_txt)
                                if fm: result[field] = f"{field}:{fm.group(1).strip()[:100]}"
                            parts = [v for v in result.values() if v]
                            paper["页面头部信息"] = "\n---\n".join(parts) if parts else d_txt[:500]
                            paper["来源"] = "知网详情页"
                        close_tab(d_tab)
                if not paper.get("来源"):
                    meta = metadata[i] if i < len(metadata) else ""
                    paper["页面头部信息"] = meta[:500]
                    paper["来源"] = "XHR元数据"

            tag = "HTML" if is_html else ("详情" if paper.get("来源")=="知网详情页" else "META")
            print(f"    [{idx}/{total}] {tag} {paper['论文标题'][:50]}")

            papers.append(paper)
            time.sleep(0.3)

        except Exception as e:
            import traceback
            err = traceback.format_exc()
            print(f"    [{idx}] ERROR: {e}")
            print(f"    {err[-200:]}")
            paper = {"序号": idx, "论文标题": titles[i] if i < len(titles) else "",
                     "作者简介原文": "", "论文元数据": metadata[i] if i < len(metadata) else "",
                     "页面头部信息": f"采集失败: {str(e)}", "来源": f"失败:{str(e)[:50]}"}
            papers.append(paper)
            continue

    return papers

def save_to_excel(papers, author_name):
    """保存到Excel"""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    df = pd.DataFrame(papers)
    ts = time.strftime("%H%M%S")
    path = os.path.join(OUTPUT_DIR, f"{author_name}_论文采集_{ts}.xlsx")
    df.to_excel(path, index=False)
    print(f"\n已保存 {len(papers)} 篇论文 → {path}")
    return path

def collect(author_name, max_papers=20):
    """主流程"""
    print(f"\n{'='*50}")
    print(f"论文采集: {author_name}")
    print(f"{'='*50}")

    # 1. 找知网tab
    print("1. 连接知网...")
    tab = find_or_create_cnki_tab()
    print(f"   Tab: {tab[:20]}")

    # 2. 搜索
    print(f"2. 搜索 {author_name}...")
    result = search_author(tab, author_name)
    print(f"   结果: {result.get('count',0)} 条 (当前第1页)")

    if int(result.get("count", 0)) == 0:
        print("   未找到结果!")
        return None

    # 3. 采集论文(HTML+元数据)
    print(f"3. 全量采集论文...")
    papers = collect_all_papers(tab, author_name, result.get("titles",[]), result.get("detail_urls",[]), max_papers)

    # 4. 保存Excel
    print(f"4. 保存Excel...")
    path = save_to_excel(papers, author_name)
    return path

if __name__ == "__main__":
    if len(sys.argv) < 2:
        name = "张伟"  # 默认测试
    else:
        name = sys.argv[1]
    max_p = 20
    if "--max" in sys.argv:
        idx = sys.argv.index("--max")
        max_p = int(sys.argv[idx+1])

    collect(name, max_p)
