

# 文档生成引擎节配置（Section Config）重构修改文档

## 1. 问题描述
在文档生成引擎的配置模块中发现严重的硬编码（Hardcode）缺陷。
**受影响文件：** `D:\PycharmProjects\hit-paper-helper\docx_manager\docx_engine\sections_config\hit_config.json`
**具体缺陷：** 配置文件中存在 `body_start_section_break`、`body_section_breaks`、`final_section_break` 三个固定字段。这种硬编码方式限定了正文节的数量及各节的配置，导致系统缺乏动态适应性。

## 2. 修改背景/原因
本项目的核心定位是**通用文档生成引擎**。即便是针对特定目标（如 HIT 专属配置），引擎也不应且无法在配置层面预先限定正文的固定节数及每一节的具体配置。当前的硬编码逻辑严重违背了通用引擎的设计原则，必须进行配置项与读取逻辑的重构。

## 3. 修改方案设计
废弃现有的固定节配置，引入**“特殊节配置（Special Section）”**及其**“配置继承机制”**。

**3.1 统一特殊节配置**
将原有的 `body_start_section_break`、`body_section_breaks`、`final_section_break` 以及之前的 `front matter` 统一废弃，合并重构为新的配置项：`special_section_break`。

**3.2 引入配置继承机制**
引擎需实现向下继承逻辑：当定义了一个特殊节后，其后续的 $n$ 个节（在遇到下一个特殊节之前）将自动继承该特殊节的页面配置（包括页脚 Footer、页眉 Header、页面大小等）。

**3.3 数据结构规范**
`special_section_break` 的数据结构必须包含以下字段：
*   节名称 (Section Name)
*   页脚 (Footer)
*   页眉 (Header)
*   页面大小 (Page Size)
*   页码重置信息 (restart_page_number)

## 4. 具体修改任务 (Action Items)

*   **任务一：重构配置文件 (`hit_config.json`)**
    *   移除所有旧版的正文节硬编码字段及 `front matter`。
    *   新增 `special_section_break` 结构。
    *   根据新的继承机制，在 HIT 专属配置中，**仅需保留“摘要”和“目录”两个特殊节配置**即可满足需求。

*   **任务二：修改读取逻辑 (`user_data_generator.py`)**
    *   同步修改该文件中关于 `hit_config.json` 的读取与解析代码。
    *   适配新的 `special_section_break` 数据结构，并确保配置继承逻辑能够被正确解析和向下传递。