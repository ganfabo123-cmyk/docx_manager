"""
style_normalizer.py — 段落样式归一化处理器 (v3)

根据 REVIEW.md 的方案实现:
1. 样式扁平化处理 (Flattening) - 递归读取 basedOn 链条,计算最终生效属性
2. 属性指纹提取 (Fingerprinting) - 提取关键属性,过滤噪音
3. 聚类与重映射 (Re-mapping) - 归一化相似指纹,创建标准样式
4. 生成 unified_style.json - 输出归一化后的样式格式
"""

import json
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path
from typing import Any, Optional
import traceback

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
TEMPLATE_DIR = BASE_DIR / "template"
STYLES_PATH = TEMPLATE_DIR / "word" / "styles.xml"
EXTRACTION_PATH = DATA_DIR / "extraction.json"
OUTPUT_PATH = DATA_DIR / "unified_style.json"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

NOISE_TAGS = {"rsid", "hint", "uiPriority", "qFormat", "kern"}

KEY_PPR_ATTRS = ["firstLine", "line", "jc", "before", "after", "left", "right"]
KEY_RPR_ATTRS = ["font_cn", "font_en", "size", "bold", "italic"]


def _local(tag: str) -> str:
    return tag.split("}", 1)[1] if "}" in tag else tag


def _parse(xml_str: Optional[str]) -> Optional[ET.Element]:
    if not xml_str:
        return None
    try:
        return ET.fromstring(xml_str)
    except ET.ParseError:
        return None


def _get_attr(el: ET.Element, attr: str, ns: str = W) -> Optional[str]:
    return el.get(f"{{{ns}}}{attr}") or el.get(attr)


class StyleResolver:
    """
    样式解析器 - 递归读取 styles.xml 的 basedOn 链条,计算最终生效属性
    
    计算公式: 最终属性 = DocDefaults + ParentStyle(basedOn) + CurrentStyle + DirectPPr + DirectRPr
    """
    
    def __init__(self, styles_path: Path):
        self.styles_path = styles_path
        self.styles: dict[str, dict] = {}
        self.doc_defaults: dict = {"pPr": {}, "rPr": {}}
        self._load_styles()
    
    def _load_styles(self):
        try:
            tree = ET.parse(str(self.styles_path))
            root = tree.getroot()
            
            for elem in root:
                local = _local(elem.tag)
                
                if local == "docDefaults":
                    self._parse_doc_defaults(elem)
                elif local == "style":
                    self._parse_style(elem)
        except Exception as e:
            print(f"Error loading styles: {e}")
            traceback.print_exc()
    
    def _parse_doc_defaults(self, elem: ET.Element):
        for child in elem:
            local = _local(child.tag)
            if local == "pPrDefault":
                pPr = child.find(f"{{{W}}}pPr")
                if pPr is not None:
                    self.doc_defaults["pPr"] = self._extract_props(pPr)
            elif local == "rPrDefault":
                rPr = child.find(f"{{{W}}}rPr")
                if rPr is not None:
                    self.doc_defaults["rPr"] = self._extract_props(rPr)
    
    def _parse_style(self, elem: ET.Element):
        style_id = _get_attr(elem, "styleId")
        if not style_id:
            return
        
        style_name = ""
        based_on = None
        
        for child in elem:
            local = _local(child.tag)
            if local == "name":
                style_name = _get_attr(child, "val") or ""
            elif local == "basedOn":
                based_on = _get_attr(child, "val")
            elif local == "pPr":
                pPr = self._extract_props(child)
                self.styles[style_id] = self.styles.get(style_id, {})
                self.styles[style_id]["pPr"] = pPr
            elif local == "rPr":
                rPr = self._extract_props(child)
                self.styles[style_id] = self.styles.get(style_id, {})
                self.styles[style_id]["rPr"] = rPr
        
        self.styles[style_id] = {
            "name": style_name,
            "basedOn": based_on,
            "pPr": self.styles.get(style_id, {}).get("pPr", {}),
            "rPr": self.styles.get(style_id, {}).get("rPr", {})
        }
    
    def _extract_props(self, elem: ET.Element) -> dict:
        props = {}
        for child in elem:
            local = _local(child.tag)
            if local in NOISE_TAGS:
                continue
            
            attrs = {}
            for k, v in child.attrib.items():
                attr_name = _local(k)
                attrs[attr_name] = v
            
            if local == "rFonts":
                props["rFonts"] = self._normalize_rfonts(attrs)
            elif local == "spacing":
                props["spacing"] = attrs
            elif local == "ind":
                props["ind"] = attrs
            elif local == "jc":
                props["jc"] = attrs.get("val", "")
            elif local in ["sz", "szCs"]:
                val = attrs.get("val")
                if not val:
                    continue
                size_val = int(val) // 2
                
                if local == "sz":
                    # sz 是基础字号，优先记录
                    props["sz"] = size_val
                    props["_sz_source"] = "sz"
                elif local == "szCs":
                    # szCs 是复杂脚本字号，仅当 sz 不存在时才作为备选
                    if "sz" not in props:
                        props["sz"] = size_val
                        props["_sz_source"] = "szCs"
            elif local in ["b", "i", "u"]:
                props[local] = attrs.get("val", "1")
            else:
                props[local] = attrs
        
        return props
    
    def _normalize_rfonts(self, attrs: dict) -> dict:
        font_en = attrs.get("ascii") or attrs.get("hAnsi") or "Times New Roman"
        font_cn = attrs.get("eastAsia") or "宋体"
        return {
            "font_en": font_en,
            "font_cn": font_cn
        }
    
    def resolve_style(self, style_id: Optional[str]) -> dict:
        if not style_id:
            return {
                "pPr": dict(self.doc_defaults["pPr"]),
                "rPr": dict(self.doc_defaults["rPr"])
            }
        
        resolved = {
            "pPr": dict(self.doc_defaults["pPr"]),
            "rPr": dict(self.doc_defaults["rPr"])
        }
        
        chain = self._get_style_chain(style_id)
        
        for sid in reversed(chain):
            if sid in self.styles:
                style = self.styles[sid]
                if style.get("pPr"):
                    resolved["pPr"].update(style["pPr"])
                if style.get("rPr"):
                    resolved["rPr"].update(style["rPr"])
        
        return resolved
    
    def _get_style_chain(self, style_id: str) -> list[str]:
        chain = []
        visited = set()
        current = style_id
        
        while current and current not in visited:
            visited.add(current)
            chain.append(current)
            
            if current in self.styles:
                current = self.styles[current].get("basedOn")
            else:
                break
        
        return chain


class FingerprintExtractor:
    """
    指纹提取器 - 将段落属性转换为标准化的指纹字典
    
    关键属性组:
    - 排版组: firstLine, line, jc, before, after
    - 字体组: font_cn, font_en, size, bold, italic
    """
    
    @staticmethod
    def extract_fingerprint(pPr_xml: Optional[str], rPr_xml: Optional[str]) -> dict:
        fingerprint = {
            "layout": {},
            "font": {}
        }
        
        if pPr_xml:
            pPr = _parse(pPr_xml)
            if pPr is not None:
                fingerprint["layout"] = FingerprintExtractor._extract_layout(pPr)
        
        if rPr_xml:
            rPr = _parse(rPr_xml)
            if rPr is not None:
                fingerprint["font"] = FingerprintExtractor._extract_font(rPr)
        
        return fingerprint
    
    @staticmethod
    def _extract_layout(pPr: ET.Element) -> dict:
        layout = {}
        
        for child in pPr:
            local = _local(child.tag)
            if local in NOISE_TAGS:
                continue
            
            if local == "spacing":
                attrs = {_local(k): v for k, v in child.attrib.items()}
                if "line" in attrs:
                    line_val = int(attrs["line"])
                    layout["line"] = line_val // 2
                if "before" in attrs:
                    layout["before"] = int(attrs["before"])
                if "after" in attrs:
                    layout["after"] = int(attrs["after"])
            elif local == "ind":
                attrs = {_local(k): v for k, v in child.attrib.items()}
                if "firstLine" in attrs:
                    layout["firstLine"] = int(attrs["firstLine"])
                if "left" in attrs:
                    layout["left"] = int(attrs["left"])
            elif local == "jc":
                val = _get_attr(child, "val")
                if val:
                    layout["jc"] = val
        
        return layout
    
    @staticmethod
    def _extract_font(rPr: ET.Element) -> dict:
        """提取字体属性，只记录有具体值的属性，并标记来源"""
        font = {}
        
        for child in rPr:
            local = _local(child.tag)
            if local in NOISE_TAGS:
                continue
            
            if local == "rFonts":
                attrs = {_local(k): v for k, v in child.attrib.items()}
                
                # 🔑 关键：只记录有具体值的属性
                if "eastAsia" in attrs and attrs["eastAsia"].strip():
                    font["font_cn"] = attrs["eastAsia"]
                    font["_src_cn"] = "direct"  # 标记: 直接格式设置
                # 注意：如果只有 hint 或 eastAsia 为空 → 不设置 font_cn
                
                if "ascii" in attrs and attrs["ascii"].strip():
                    font["font_en"] = attrs["ascii"]
                    font["_src_en"] = "direct"
                elif "hAnsi" in attrs and attrs["hAnsi"].strip():
                    font["font_en"] = attrs["hAnsi"]
                    font["_src_en"] = "direct"
                    
            elif local == "sz":
                val = _get_attr(child, "val")
                if val and val.strip():
                    font["size"] = int(val) // 2
                    font["_src_size"] = "direct"
                    
            elif local == "b":
                val = _get_attr(child, "val", W)
                if val and val not in ("0", "false"):
                    font["bold"] = True
                    font["_src_bold"] = "direct"
                    
            elif local == "i":
                val = _get_attr(child, "val", W)
                if val and val not in ("0", "false"):
                    font["italic"] = True
                    font["_src_italic"] = "direct"
        
        return font


def _merge_font_properties(direct_font: dict, resolved_rPr: dict) -> dict:
    """
    属性级继承合并：只有当 direct 未显式设置时，才从 resolved 继承
    
    优先级: direct > style_definition > basedOn_chain > docDefaults
    """
    merged = dict(direct_font)  # 以直接格式为基底
    
    # 🎯 中文字体继承
    if "font_cn" not in merged or merged.get("_src_cn") != "direct":
        # direct 未设置 → 从 resolved 继承
        resolved_fonts = resolved_rPr.get("rFonts", {})
        if resolved_fonts.get("font_cn"):
            merged["font_cn"] = resolved_fonts["font_cn"]
            merged["_src_cn"] = "inherited"
    
    # 🎯 英文字体继承
    if "font_en" not in merged or merged.get("_src_en") != "direct":
        resolved_fonts = resolved_rPr.get("rFonts", {})
        # 优先 ascii, 其次 hAnsi
        en_font = resolved_fonts.get("font_en")
        if en_font:
            merged["font_en"] = en_font
            merged["_src_en"] = "inherited"
    
    # 🔑 字号继承：优先 sz，其次 szCs（容错机制）
    if "size" not in merged or merged.get("_src_size") != "direct":
        # 优先取 resolved 的 sz，其次取 szCs（如果存在）
        resolved_size = resolved_rPr.get("sz") or resolved_rPr.get("szCs")
        if resolved_size is not None:
            # resolved_size 已经是 //2 后的值（在 _extract_props 中处理过）
            merged["size"] = resolved_size
            merged["_src_size"] = "inherited"
    
    # 🎯 修饰符继承 (bold/italic)
    for attr in ["bold", "italic"]:
        if attr not in merged or merged.get(f"_src_{attr}") != "direct":
            if attr in resolved_rPr:
                merged[attr] = resolved_rPr[attr]
                merged[f"_src_{attr}"] = "inherited"
    
    return merged


def _merge_layout_properties(direct_layout: dict, resolved_pPr: dict) -> dict:
    """
    布局属性级继承
    """
    merged = dict(direct_layout)
    
    # 段前间距
    if "before" not in merged:
        spacing = resolved_pPr.get("spacing", {})
        if "before" in spacing:
            merged["before"] = int(spacing["before"])
    
    # 段后间距
    if "after" not in merged:
        spacing = resolved_pPr.get("spacing", {})
        if "after" in spacing:
            merged["after"] = int(spacing["after"])
    
    # 行距
    if "line" not in merged:
        spacing = resolved_pPr.get("spacing", {})
        if "line" in spacing:
            merged["line"] = int(spacing["line"]) // 2
    
    # 对齐方式
    if "jc" not in merged and "jc" in resolved_pPr:
        merged["jc"] = resolved_pPr["jc"]
    
    # 首行缩进
    if "firstLine" not in merged:
        ind = resolved_pPr.get("ind", {})
        if "firstLine" in ind:
            merged["firstLine"] = int(ind["firstLine"])
    
    return merged


def _clean_fingerprint_sources(fingerprint: dict) -> dict:
    """
    清理内部标记字段，输出前去掉 _src_* 和 _sz_source 标记
    """
    cleaned = {}
    if "layout" in fingerprint:
        cleaned["layout"] = {k: v for k, v in fingerprint["layout"].items() 
                            if not k.startswith("_src_") and k != "_sz_source"}
    if "font" in fingerprint:
        cleaned["font"] = {k: v for k, v in fingerprint["font"].items() 
                          if not k.startswith("_src_") and k != "_sz_source"}
    return cleaned


class StyleCluster:
    """
    样式聚类器 - 将相似指纹归类,创建标准样式 (基于阈值动态聚类)
    """
    
    def __init__(self, threshold=5):
        # 核心数据结构改为列表，每个元素是一个簇：{"fingerprint": dict, "items": list}
        self.clusters: list[dict] = []
        self.threshold = threshold
        
    def add_fingerprint(self, fingerprint: dict, para_index: int, style_id: Optional[str],text: str = ""):
        item = {
            "index": para_index, 
            "original_style": style_id,
            "text": text
        }
        
        # 1. 遍历现有簇，看是否能塞进某个簇
        for cluster in self.clusters:
            if self._is_similar(cluster["fingerprint"], fingerprint):
                cluster["items"].append(item)
                return
        
        # 2. 如果找不到相似簇，自立门户，以当前指纹作为该簇的“质心”
        self.clusters.append({
            "fingerprint": fingerprint,
            "items": [item]
        })

    def _is_similar(self, f1: dict, f2: dict) -> bool:
        l1, l2 = f1.get("layout", {}), f2.get("layout", {})
        
        # 1. 缩进比较 (容差放大到 20)
        # 解决 498 和 507 的历史遗留分歧！
        # 把缩进除以 240 取整，直接转换为“缩进了几个字符”
        chars1 = round(l1.get("firstLine", 0) / 240)
        chars2 = round(l2.get("firstLine", 0) / 240)
        if chars1 != chars2:
            return False
            
        # 2. 行距比较 (容差 10)
        # 缺失的默认值为 0，有值的如 150，绝对会正确地分离开
        line1 = l1.get("line", 0)
        line2 = l2.get("line", 0)
        if abs(line1 - line2) > 10:
            return False
            
        # 3. 对齐比较
        if l1.get("jc", "both") != l2.get("jc", "both"):
            return False
            
        # 4. 字体比较：坚决忽略 font_en！(破除输入法干扰)
        # 只比较中文字体
        font1, font2 = f1.get("font", {}), f2.get("font", {})
        if font1.get("font_cn", "宋体") != font2.get("font_cn", "宋体"):
            return False
            
        # 🔑 5. 字号比较：严格相等（标题字号差异必须保留）
        # 解决 14pt/15pt 标题被错误合并的问题
        sz1 = font1.get("size", 24)
        sz2 = font2.get("size", 24)
        if sz1 != sz2:  # 改为严格相等，而不是容差 1
            return False
            
        # 6. 修饰符比较
        if font1.get("bold", False) != font2.get("bold", False):
            return False
        if font1.get("italic", False) != font2.get("italic", False):
            return False
            
        return True
        
    def _generate_style_name(self, index: int, fingerprint: dict) -> str:
        """为生成的簇起一个有可读性的名字，方便你 debug 检查"""
        font_sz = fingerprint.get("font", {}).get("size", 24)
        jc = fingerprint.get("layout", {}).get("jc", "both")
        is_bold = "_Bold" if fingerprint.get("font", {}).get("bold") else ""
        return f"Cluster_{index}_sz{font_sz}_{jc}{is_bold}"
    
    def get_standard_styles(self) -> dict[str, dict]:
        standard_styles = {}
        
        # 按簇内包含的段落数量从大到小排序，这样最常见的样式(如正文)会被排在前面
        sorted_clusters = sorted(self.clusters, key=lambda c: len(c["items"]), reverse=True)
        
        for i, cluster in enumerate(sorted_clusters, 1):
            items = cluster["items"]
            representative = cluster["fingerprint"]
            
            # 生成全局唯一的 Key，如 Style_001
            cluster_key = f"Style_{i:03d}"  
            style_name = self._generate_style_name(i, representative)
            
            standard_styles[cluster_key] = {
                "style_name": style_name,
                "fingerprint": representative,
                "occurrence_count": len(items),
                "paragraphs": items
            }
        
        return standard_styles


class ClusterFilter:
    """
    簇筛选器 - 论文排版助手的"审查与收编"系统
    
    三大法则:
    1. 频次斩杀线 - occurrence_count >= 3 保留
    2. 语义免死金牌 - 识别特殊标题和封面内容
    3. 黑名单过滤 - 过滤异常字号和纯空白段落
    
    强行收编:
    将不合规簇合并到距离最近的合规簇
    """
    
    SEMANTIC_PATTERNS = [
        r"^摘\s*要$",
        r"^Abstract$",
        r"^目\s*录$",
        r"^参考文献$",
        r"^致\s*谢$",
        r"^第[一二三四五六七八九十\d]+章",
        r"^\d+\.\d+",
        r"哈尔滨工业大学",
        r"本科.*设计.*论文"
    ]
    
    def __init__(self, standard_styles: dict[str, dict], body_elements: list[dict]):
        self.styles = standard_styles
        self.body_elements = body_elements
        self.valid_clusters: dict[str, dict] = {}
        self.garbage_clusters: dict[str, dict] = {}
    
    def filter_and_merge(self) -> dict[str, dict]:
        """
        执行筛选与强行收编流程
        
        Returns:
            清洗后的合规簇字典
        """
        self._screen_clusters()
        self._force_merge_garbage()
        
        return self.valid_clusters
    
    def _screen_clusters(self):
        """筛选审查 - 将簇分为合规与垃圾两类"""
        for cluster_key, cluster_data in self.styles.items():
            is_valid = self._check_cluster_validity(cluster_data)
            
            if is_valid:
                self.valid_clusters[cluster_key] = cluster_data
            else:
                self.garbage_clusters[cluster_key] = cluster_data
        
        print(f"\n筛选完毕：保留合规簇 {len(self.valid_clusters)} 个，拦截垃圾簇 {len(self.garbage_clusters)} 个")
    
    def _check_cluster_validity(self, cluster_data: dict) -> bool:
        """
        检查簇是否符合保留标准
        
        Args:
            cluster_data: 簇数据
            
        Returns:
            是否合规
        """
        fingerprint = cluster_data.get("fingerprint", {})
        
        if self._is_blacklisted(fingerprint):
            return False
        
        if cluster_data["occurrence_count"] >= 3:
            return True
        
        if self._has_semantic_immunity(cluster_data):
            return True
        
        return False
    
    def _is_blacklisted(self, fingerprint: dict) -> bool:
        """黑名单过滤 - 检查是否为异常样式"""
        font = fingerprint.get("font", {})
        size = font.get("size", 24)
        
        if size < 8 or size > 50:
            return True
        
        return False
    
    def _has_semantic_immunity(self, cluster_data: dict) -> bool:
        """语义免死金牌 - 检查是否为特殊标题或封面"""
        paragraphs = cluster_data.get("paragraphs", [])
        
        for para_info in paragraphs[:3]:
            para_index = para_info.get("index")
            text = self._get_paragraph_text(para_index)
            
            if not text:
                continue
            
            if any(self._match_pattern(text, pattern) for pattern in self.SEMANTIC_PATTERNS):
                return True
        
        if any(p.get("index", 999) < 20 for p in paragraphs):
            return True
        
        return False
    
    def _get_paragraph_text(self, para_index: int) -> str:
        """根据段落索引获取文本内容"""
        for elem in self.body_elements:
            if elem.get("index") == para_index:
                return elem.get("text", "")
        return ""
    
    def _match_pattern(self, text: str, pattern: str) -> bool:
        """匹配正则模式"""
        import re
        try:
            return bool(re.search(pattern, text))
        except Exception:
            return False
    
    def _force_merge_garbage(self):
        """强行收编 - 将垃圾簇合并到最近的合规簇"""
        if not self.valid_clusters:
            print("警告：没有合规簇，无法执行收编！")
            return
        
        for garbage_key, garbage_data in self.garbage_clusters.items():
            closest_key = self._find_closest_valid_cluster(garbage_data)
            
            if closest_key:
                self.valid_clusters[closest_key]["paragraphs"].extend(
                    garbage_data["paragraphs"]
                )
                self.valid_clusters[closest_key]["occurrence_count"] += \
                    garbage_data["occurrence_count"]
    
    def _find_closest_valid_cluster(self, garbage_data: dict) -> Optional[str]:
        """寻找距离最近的合规簇"""
        garbage_fingerprint = garbage_data.get("fingerprint", {})
        closest_key = None
        min_distance = float('inf')
        
        for valid_key, valid_data in self.valid_clusters.items():
            distance = self._calculate_style_distance(
                garbage_fingerprint,
                valid_data.get("fingerprint", {})
            )
            
            if distance < min_distance:
                min_distance = distance
                closest_key = valid_key
        
        return closest_key
    
    def _calculate_style_distance(self, f1: dict, f2: dict) -> float:
        """
        计算两个样式指纹之间的距离
        
        距离越小表示越相似
        字号权重最高，行距次之，其他属性权重较低
        """
        distance = 0.0
        
        font1 = f1.get("font", {})
        font2 = f2.get("font", {})
        
        size1 = font1.get("size", 24)
        size2 = font2.get("size", 24)
        distance += abs(size1 - size2) * 10
        
        if font1.get("font_cn", "宋体") != font2.get("font_cn", "宋体"):
            distance += 50
        
        if font1.get("bold", False) != font2.get("bold", False):
            distance += 30
        
        layout1 = f1.get("layout", {})
        layout2 = f2.get("layout", {})
        
        line1 = layout1.get("line", 150)
        line2 = layout2.get("line", 150)
        distance += abs(line1 - line2) * 0.5
        
        if layout1.get("jc", "both") != layout2.get("jc", "both"):
            distance += 20
        
        return distance
    
    def get_merge_report(self) -> dict:
        """生成合并报告"""
        return {
            "total_valid": len(self.valid_clusters),
            "total_garbage": len(self.garbage_clusters),
            "garbage_details": [
                {
                    "style_name": data.get("style_name", ""),
                    "occurrence_count": data.get("occurrence_count", 0),
                    "merged_into": self._find_closest_valid_cluster(data)
                }
                for data in self.garbage_clusters.values()
            ]
        }


class StyleAnalyzer:
    """
    样式分析器 - 对聚类结果进行深度分析
    
    主要功能:
    1. 分析字体字号相同的簇 - 识别可能需要合并的簇
    2. 分析单例簇 - 识别可能需要特殊处理的孤立段落
    3. 生成统计报告 - 提供决策支持
    """
    
    def __init__(self, standard_styles: dict[str, dict]):
        self.styles = standard_styles
        self.font_size_groups: dict[str, list] = defaultdict(list)
        self.singleton_clusters: list[dict] = []
    
    def analyze(self) -> dict:
        """
        执行完整分析流程
        
        Returns:
            包含所有分析结果的字典
        """
        self._analyze_font_size_groups()
        self._analyze_singleton_clusters()
        
        return {
            "font_size_analysis": self._get_font_size_report(),
            "singleton_analysis": self._get_singleton_report(),
            "summary": self._get_summary()
        }
    
    def _analyze_font_size_groups(self):
        """按字体+字号分组，识别可能需要合并的簇"""
        for style_key, style_data in self.styles.items():
            font_info = style_data.get("fingerprint", {}).get("font", {})
            
            font_cn = font_info.get("font_cn", "宋体")
            font_en = font_info.get("font_en", "Times New Roman")
            size = font_info.get("size", 24)
            
            group_key = f"{font_cn}|{font_en}|{size}pt"
            
            cluster_info = {
                "style_key": style_key,
                "style_name": style_data.get("style_name", ""),
                "occurrence_count": style_data.get("occurrence_count", 0),
                "layout_diff": self._extract_layout_signature(style_data)
            }
            
            existing_keys = [c["style_key"] for c in self.font_size_groups[group_key]]
            if style_key not in existing_keys:
                self.font_size_groups[group_key].append(cluster_info)
    
    def _extract_layout_signature(self, style_data: dict) -> str:
        """提取布局特征签名，用于识别相同字体字号但布局不同的簇"""
        layout = style_data.get("fingerprint", {}).get("layout", {})
        
        parts = []
        if "firstLine" in layout:
            parts.append(f"缩进{layout['firstLine']}")
        if "line" in layout:
            parts.append(f"行距{layout['line']}")
        if "jc" in layout:
            parts.append(f"对齐{layout['jc']}")
        if "before" in layout:
            parts.append(f"段前{layout['before']}")
        if "after" in layout:
            parts.append(f"段后{layout['after']}")
        
        return " | ".join(parts) if parts else "默认布局"
    
    def _analyze_singleton_clusters(self):
        """识别只包含一个段落的簇（单例簇）"""
        for style_key, style_data in self.styles.items():
            count = style_data.get("occurrence_count", 0)
            if count == 1:
                para_info = style_data.get("paragraphs", [{}])[0]
                
                self.singleton_clusters.append({
                    "style_key": style_key,
                    "style_name": style_data.get("style_name", ""),
                    "para_index": para_info.get("index"),
                    "original_style": para_info.get("original_style"),
                    "fingerprint": style_data.get("fingerprint", {})
                })
    
    def _get_font_size_report(self) -> dict:
        """生成字体字号分组报告"""
        report = {
            "total_groups": len(self.font_size_groups),
            "groups_with_multiple_clusters": 0,
            "details": []
        }
        
        sorted_groups = sorted(
            self.font_size_groups.items(),
            key=lambda x: sum(c["occurrence_count"] for c in x[1]),
            reverse=True
        )
        
        for group_key, clusters in sorted_groups:
            if len(clusters) > 1:
                report["groups_with_multiple_clusters"] += 1
                
                total_occurrences = sum(c["occurrence_count"] for c in clusters)
                
                sorted_clusters = sorted(clusters, key=lambda x: x["occurrence_count"], reverse=True)
                
                group_detail = {
                    "font_signature": group_key,
                    "cluster_count": len(clusters),
                    "total_occurrences": total_occurrences,
                    "clusters": sorted_clusters
                }
                
                report["details"].append(group_detail)
        
        return report
    
    def _get_singleton_report(self) -> dict:
        """生成单例簇报告"""
        return {
            "total_singletons": len(self.singleton_clusters),
            "singletons": sorted(self.singleton_clusters, key=lambda x: x["style_key"])
        }
    
    def _get_summary(self) -> dict:
        """生成总体摘要"""
        total_clusters = len(self.styles)
        total_paragraphs = sum(s.get("occurrence_count", 0) for s in self.styles.values())
        
        multi_cluster_groups = sum(
            1 for clusters in self.font_size_groups.values() if len(clusters) > 1
        )
        
        return {
            "total_clusters": total_clusters,
            "total_paragraphs": total_paragraphs,
            "total_font_size_groups": len(self.font_size_groups),
            "multi_cluster_groups": multi_cluster_groups,
            "singleton_clusters": len(self.singleton_clusters),
            "merge_candidates": multi_cluster_groups,
            "optimization_potential": f"{(multi_cluster_groups / max(total_clusters, 1) * 100):.1f}%"
        }
    
    def print_report(self):
        """打印详细的分析报告"""
        analysis = self.analyze()
        
        print("\n" + "=" * 70)
        print("样式聚类分析报告")
        print("=" * 70)
        
        summary = analysis["summary"]
        print("\n【总体统计】")
        print(f"  总簇数: {summary['total_clusters']}")
        print(f"  总段落数: {summary['total_paragraphs']}")
        print(f"  字体字号组合数: {summary['total_font_size_groups']}")
        print(f"  多簇组合数: {summary['multi_cluster_groups']}")
        print(f"  单例簇数: {summary['singleton_clusters']}")
        print(f"  潜在优化空间: {summary['optimization_potential']}")
        
        font_report = analysis["font_size_analysis"]
        print("\n【字体字号相同的簇】")
        print(f"  发现 {font_report['groups_with_multiple_clusters']} 组字体字号相同但布局不同的簇:")
        
        for i, group in enumerate(font_report["details"][:10], 1):
            print(f"\n  {i}. {group['font_signature']}")
            print(f"     包含 {group['cluster_count']} 个簇, 共 {group['total_occurrences']} 个段落")
            
            for cluster in group["clusters"][:3]:
                print(f"       - {cluster['style_name']}: {cluster['occurrence_count']} 次 [{cluster['layout_diff']}]")
        
        singleton_report = analysis["singleton_analysis"]
        print("\n【单例簇（只出现1次的段落）】")
        print(f"  发现 {singleton_report['total_singletons']} 个单例簇:")
        
        for i, singleton in enumerate(singleton_report["singletons"][:10], 1):
            para_idx = singleton["para_index"]
            orig_style = singleton.get("original_style", "无")
            print(f"  {i}. 段落 {para_idx} (原样式: {orig_style})")
            print(f"     {singleton['style_name']}")
        
        if singleton_report["total_singletons"] > 10:
            print(f"  ... 还有 {singleton_report['total_singletons'] - 10} 个单例簇未显示")

def _is_default_font(font_dict: dict) -> bool:
    """判断字体字典是否仅包含默认值（无实际格式设置）"""
    if not font_dict:
        return True
    defaults = {
        "font_cn": "宋体", 
        "font_en": "Times New Roman", 
        "size": 24,  # 注意：代码中 size 已 //2，所以默认是 24
        "bold": False,
        "italic": False
    }
    # 所有存在的键都等于默认值 → 视为"无实际格式"
    return all(font_dict.get(k) == v for k, v in defaults.items() if k in font_dict)

def normalize_document():
    """
    主函数 - 执行文档样式归一化流程
    """
    print("=" * 70)
    print("文档样式归一化处理器 v3")
    print("=" * 70)
    
    print("\n[1/4] 加载样式定义...")
    resolver = StyleResolver(STYLES_PATH)
    print(f"  - 已加载 {len(resolver.styles)} 个样式定义")
    print(f"  - 文档默认样式: {resolver.doc_defaults}")
    
    print("\n[2/4] 加载文档内容...")
    with open(EXTRACTION_PATH, encoding="utf-8") as f:
        extraction_data = json.load(f)
    
    body_elements = extraction_data.get("body_elements", [])
    print(f"  - 已加载 {len(body_elements)} 个文档元素")
    
    print("\n[3/4] 提取指纹并聚类...")
    cluster = StyleCluster()
    for elem in body_elements:
        if elem.get("type") != "paragraph":
            continue
        
        para_index = elem.get("index")
        style_id = elem.get("style")
        text = elem.get("text", "").strip()
        if text == '摘  要':
            pass
        if not text:
            fingerprint = {
                "layout": {"line": 150},
                "font": {"size": 12, "font_cn": "宋体", "font_en": "Times New Roman"}
            }
            cluster.add_fingerprint(fingerprint=fingerprint, para_index=para_index,style_id= style_id,text = text)
            continue
            
        pPr_xml = elem.get("pPr")
        resolved = resolver.resolve_style(style_id)
        
        rPr_in_pPr = None
        if pPr_xml:
            pPr_elem = _parse(pPr_xml)
            if pPr_elem is not None:
                rPr_in_pPr_elem = pPr_elem.find(f"{{{W}}}rPr")
                if rPr_in_pPr_elem is not None:
                    rPr_in_pPr = ET.tostring(rPr_in_pPr_elem, encoding="unicode")
        
        fingerprint = FingerprintExtractor.extract_fingerprint(pPr_xml, rPr_in_pPr)
        
        # 🔑 属性级合并（核心改动）
        if resolved.get("rPr"):
            fingerprint["font"] = _merge_font_properties(
                direct_font=fingerprint.get("font", {}),
                resolved_rPr=resolved["rPr"]
            )
        
        # 布局属性同理
        if resolved.get("pPr"):
            fingerprint["layout"] = _merge_layout_properties(
                direct_layout=fingerprint.get("layout", {}),
                resolved_pPr=resolved["pPr"]
            )
        
        # 清理内部标记字段，输出前去掉 _src_* 标记
        fingerprint = _clean_fingerprint_sources(fingerprint)
        
        cluster.add_fingerprint(fingerprint, para_index, style_id, text=text)
    standard_styles = cluster.get_standard_styles()
    print(f"  - 识别出 {len(standard_styles)} 个标准样式簇")
    
    print("\n[3.5/4] 执行簇筛选与强行收编...")
    filter_engine = ClusterFilter(standard_styles, body_elements)
    cleaned_styles = filter_engine.filter_and_merge()
    
    merge_report = filter_engine.get_merge_report()
    print(f"  - 合并报告：")
    print(f"    合规簇: {merge_report['total_valid']} 个")
    print(f"    垃圾簇: {merge_report['total_garbage']} 个 (已强行收编)")
    
    standard_styles = cleaned_styles
    
    print("\n[4/4] 生成归一化样式文件...")
    
    unified_data = {
        "metadata": {
            "source": str(EXTRACTION_PATH),
            "total_paragraphs": len([e for e in body_elements if e.get("type") == "paragraph"]),
            "total_clusters": len(standard_styles)
        },
        "standard_styles": standard_styles,
        "style_definitions": {}
    }
    
    for style_id, style_info in resolver.styles.items():
        unified_data["style_definitions"][style_id] = {
            "name": style_info.get("name", ""),
            "basedOn": style_info.get("basedOn"),
            "resolved_pPr": resolver.resolve_style(style_id)["pPr"],
            "resolved_rPr": resolver.resolve_style(style_id)["rPr"]
        }
    
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(unified_data, f, ensure_ascii=False, indent=2)
    
    print(f"  - 已写入: {OUTPUT_PATH}")
    
    print("\n" + "=" * 70)
    print("样式归一化完成!")
    print("=" * 70)
    
    print("\n统计信息:")
    print(f"  - 原始样式数量: {len(resolver.styles)}")
    print(f"  - 归一化后簇数量: {len(standard_styles)}")
    print(f"  - 压缩比: {len(standard_styles) / max(len(resolver.styles), 1) * 100:.1f}%")
    
    print("\n前5个最大簇:")
    sorted_clusters = sorted(standard_styles.items(), key=lambda x: x[1]["occurrence_count"], reverse=True)
    for i, (cluster_key, cluster_data) in enumerate(sorted_clusters[:5], 1):
        print(f"  {i}. {cluster_data['style_name']}: {cluster_data['occurrence_count']} 个段落")
    
    analyzer = StyleAnalyzer(standard_styles)
    analyzer.print_report()
    
    analysis = analyzer.analyze()
    analysis_output = DATA_DIR / "style_analysis.json"
    with open(analysis_output, "w", encoding="utf-8") as f:
        json.dump(analysis, f, ensure_ascii=False, indent=2)
    print(f"\n详细分析报告已保存至: {analysis_output}")
    
    merge_report_output = DATA_DIR / "merge_report.json"
    with open(merge_report_output, "w", encoding="utf-8") as f:
        json.dump(merge_report, f, ensure_ascii=False, indent=2)
    print(f"合并报告已保存至: {merge_report_output}")


if __name__ == "__main__":
    normalize_document()
