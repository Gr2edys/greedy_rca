#!/usr/bin/env python3
"""
从华为 UNC HedEx 文档包中提取所有 5GC 网元相关的 MML 命令。

用法：
    python3 extract_5gc_mml.py <文档根目录> [输出JSON路径]

示例：
    python3 extract_5gc_mml.py /mnt/c/greedy/unc_output/resources

原理：
    1. 解析 navi.xml 找到所有 MML Document 文件路径
    2. 逐个读取 HTML 文件，提取"适用NF"或"适用网元"字段
    3. 筛选出标注包含 AMF/SMF/NRF/NSSF/SMSF/UPF 的命令
    4. 输出 JSON
"""

import re
import os
import sys
import json
import time


def parse_navi_xml(navi_path):
    """从 navi.xml 提取所有 MML Document 文件路径"""
    with open(navi_path, 'r', encoding='utf-8') as f:
        content = f.read()

    doc_paths = set()
    # 匹配 url="...MML/Document/xxx.html"
    for m in re.finditer(r'url="([^"]*MML/Document/[^"]+\.html)"', content):
        path = m.group(1)
        # 处理相对路径 ../ 
        path = path.replace('../', '')
        doc_paths.add(path)

    # 也匹配 url="mml/document/xxx.html"
    for m in re.finditer(r'url="(mml/document/[^"]+\.html)"', content, re.IGNORECASE):
        doc_paths.add(m.group(1))

    return sorted(doc_paths)


def extract_nf_from_html(html_text):
    """
    从 MML 命令 HTML 中提取"适用NF"信息。
    返回 NF 字符串，如 "AMF"、"SMF、NRF"、"SGSN、MME" 等。
    如果没有标注则返回空字符串。
    """
    for m in re.finditer(r'适用(NF|网元)[：:]\s*(.+?)(?:</(?:strong|span|p)>)', html_text):
        nf_text = m.group(2).strip()
        # 清除残留 HTML 标签
        nf_clean = re.sub(r'<[^>]+>', '', nf_text).strip()
        # 统一分隔符
        nf_clean = nf_clean.replace('、', ',').replace('，', ',').replace(' ', '')
        return nf_clean
    return ''


def extract_cmd_name_from_html(html_text):
    """从 HTML <title> 中提取命令名，如 'ADD GUAMI'"""
    title_match = re.search(r'<title>(.+?)</title>', html_text)
    if not title_match:
        return ''
    title = title_match.group(1)
    cmd_match = re.search(r'[（(]([A-Z][A-Z\s]{2,})[）)]', title)
    if cmd_match:
        return cmd_match.group(1).strip()
    return ''


def main(resources_dir, output_path):
    start_time = time.time()

    # ---- 1. 定义 5GC 网元 ----
    GC_NFS = {'AMF', 'SMF', 'NRF', 'NSSF', 'SMSF', 'UPF'}

    # ---- 2. 解析 navi.xml ----
    navi_path = os.path.join(resources_dir, 'navi.xml')
    if not os.path.exists(navi_path):
        print(f"错误: 找不到 navi.xml ({navi_path})")
        sys.exit(1)

    print("解析 navi.xml ...")
    doc_paths = parse_navi_xml(navi_path)
    print(f"  找到 {len(doc_paths)} 个 MML Document 条目")

    # ---- 3. 逐个处理 MML 文件 ----
    gc_commands = []      # 5GC 命令列表
    no_nf_commands = []   # 无 NF 标注的命令（平台通用）
    stats = {'total': 0, 'gc': 0, 'no_nf': 0, 'error': 0}

    for i, rel_path in enumerate(doc_paths):
        full_path = os.path.join(resources_dir, rel_path)
        stats['total'] += 1

        if not os.path.exists(full_path):
            stats['error'] += 1
            continue

        # 只读前 5000 字节即可（NF 信息在文件头部）
        try:
            with open(full_path, 'r', encoding='gb2312', errors='ignore') as f:
                html = f.read(5000)
        except Exception:
            stats['error'] += 1
            continue

        # 提取命令名
        cmd_name = extract_cmd_name_from_html(html)
        if not cmd_name:
            # 回退：从文件名提取
            basename = os.path.basename(rel_path).replace('.html', '')
            cmd_name = basename.upper().replace('_', ' ')

        # 提取适用NF
        nf_text = extract_nf_from_html(html)

        if not nf_text:
            stats['no_nf'] += 1
            no_nf_commands.append(cmd_name)
            continue

        # 判断是否 5GC
        nf_set = set(n.strip() for n in nf_text.split(',') if n.strip())
        if nf_set & GC_NFS:
            stats['gc'] += 1
            gc_commands.append({
                'cmd': cmd_name,
                'nf': nf_text,
                'file': rel_path
            })
        # else: 纯 2G/3G/4G 命令，跳过

        # 进度提示
        if (i + 1) % 2000 == 0:
            elapsed = time.time() - start_time
            print(f"  已处理 {i+1}/{len(doc_paths)} ... ({elapsed:.1f}s)")

    # ---- 4. 统计汇总 ----
    elapsed = time.time() - start_time

    # 按 NF 统计
    nf_count = {}
    for item in gc_commands:
        for nf in item['nf'].split(','):
            nf = nf.strip()
            if nf in GC_NFS:
                nf_count[nf] = nf_count.get(nf, 0) + 1

    print(f"\n{'='*60}")
    print(f"处理完成 (耗时 {elapsed:.1f}s)")
    print(f"{'='*60}")
    print(f"总 MML 文件:             {stats['total']}")
    print(f"5GC 网元命令:            {stats['gc']}")
    print(f"平台通用命令 (无NF标注):  {stats['no_nf']}")
    print(f"读取错误:                {stats['error']}")
    print(f"\n各 5GC NF 涉及命令数 (含跨NF共享):")
    for nf in sorted(nf_count.keys()):
        print(f"  {nf:6s}: {nf_count[nf]:4d} 条")

    # ---- 5. 生成 JSON ----
    # 简洁版：仅命令名列表
    cmd_list = sorted(set(item['cmd'] for item in gc_commands))

    # 详细版：含 NF 信息和文件路径
    output = {
        'source': 'UNC 20.13.2 产品文档',
        'total_5gc_commands': len(cmd_list),
        'commands': cmd_list,
        'detail': gc_commands,
        'statistics': {
            'total_mml_files': stats['total'],
            'gc_count': stats['gc'],
            'no_nf_count': stats['no_nf'],
            'per_nf': nf_count
        }
    }

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\nJSON 已输出: {output_path}")
    print(f"  顶层字段: commands (纯命令名列表), detail (含NF/文件路径), statistics (统计)")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    resources_dir = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else './unc_5gc_mml.json'

    main(resources_dir, output_path)
