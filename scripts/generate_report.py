#!/usr/bin/env python3
"""
电销录音管理教练 - Word 报告生成器
将分析结果生成为格式化的 Word 文档。
纯 python-docx 实现。

用法：python3 generate_report.py <分析结果.json>
"""

import json
import os
import sys
from datetime import datetime

try:
    from docx import Document
    from docx.shared import Inches, Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
except ImportError:
    print("错误：需要安装 python-docx。运行：pip3 install python-docx", file=sys.stderr)
    sys.exit(1)


def set_cell_shading(cell, color):
    """设置表格单元格底色"""
    shading_elm = cell._element.get_or_add_tcPr()
    shading = shading_elm.makeelement(qn('w:shd'), {
        qn('w:val'): 'clear',
        qn('w:color'): 'auto',
        qn('w:fill'): color,
    })
    shading_elm.append(shading)


def add_styled_table(doc, headers, rows, col_widths=None):
    """添加带样式的表格"""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 表头
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(10)
                run.font.name = 'Arial'
        set_cell_shading(cell, '2B579A')
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)

    # 数据行
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            cell = table.rows[r + 1].cells[c]
            cell.text = str(value)
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
                    run.font.name = 'Arial'
            if r % 2 == 1:
                set_cell_shading(cell, 'F2F2F2')

    doc.add_paragraph()  # 表后空行
    return table


def add_bullet_paragraph(doc, text, prefix=""):
    """添加带前缀的项目符号段落"""
    p = doc.add_paragraph(f"{prefix} {text}", style='List Bullet')
    for run in p.runs:
        run.font.size = Pt(11)
        run.font.name = 'Arial'
    return p


def add_section_heading(doc, number, title):
    """添加带编号的章节标题"""
    doc.add_heading(f'{number}、{title}', level=1)


def generate_single_report(doc, data, sections):
    """生成单通录音分析报告"""
    # 一、通话摘要
    summary = sections.get('call_summary', {})
    if summary:
        add_section_heading(doc, '一', '通话摘要')
        headers = ['维度', '内容']
        rows = [
            ['客户', summary.get('customer', '未知')],
            ['需求', summary.get('need', '未明确提及')],
            ['异议', summary.get('objection', '未提及')],
            ['销售表达', summary.get('sales_pitched', '未明确提及')],
            ['结果', summary.get('outcome', '未明确')],
            ['时长', summary.get('duration', '未知')],
        ]
        add_styled_table(doc, headers, rows)
        doc.add_paragraph()

    # 二、关键标签
    tags = sections.get('tags', {})
    if tags:
        add_section_heading(doc, '二', '关键标签')
        if tags.get('core'):
            doc.add_heading('核心标签', level=3)
            for tag in tags['core']:
                add_bullet_paragraph(doc, tag, '🔹')
        if tags.get('problems'):
            doc.add_heading('问题标签', level=3)
            for tag in tags['problems']:
                add_bullet_paragraph(doc, tag, '🔴')
        if tags.get('strengths'):
            doc.add_heading('能力标签', level=3)
            for tag in tags['strengths']:
                add_bullet_paragraph(doc, tag, '✅')
        doc.add_paragraph()

    # 三、话术评分
    scores = sections.get('scores', {})
    overall = sections.get('overall_score', {})
    if scores:
        add_section_heading(doc, '三', '话术评分')
        headers = ['评分维度', '得分', '评价理由']
        dimension_names = {
            'opening': '开场有效性',
            'need_discovery': '需求挖掘充分度',
            'objection_handling': '异议处理能力',
            'value_clarity': '价值表达清晰度',
            'next_step': '推进下一步明确度',
        }
        rows = []
        for key, name in dimension_names.items():
            item = scores.get(key, {})
            score = item.get('score', '-')
            reason = item.get('reason', '未评价')
            rows.append([name, f'{score}/10', reason])
        add_styled_table(doc, headers, rows)

        if overall:
            p = doc.add_paragraph()
            run = p.add_run(f"综合得分：{overall.get('score', '-')}/10")
            run.font.size = Pt(14)
            run.bold = True
            run.font.name = 'Arial'
            if overall.get('level'):
                p2 = doc.add_paragraph()
                run2 = p2.add_run(f"等级：{overall['level']}")
                run2.font.size = Pt(11)
                run2.font.name = 'Arial'
        doc.add_paragraph()

    # 四、话术红黑榜
    script_review = sections.get('script_review', {})
    if script_review:
        add_section_heading(doc, '四', '话术红黑榜')
        if script_review.get('good_points'):
            doc.add_heading('✅ 亮点（做得好）', level=3)
            for point in script_review['good_points']:
                moment = point.get('moment', '')
                what = point.get('what', '')
                text = f"{moment}：{what}" if moment else what
                add_bullet_paragraph(doc, text, '👍')

        if script_review.get('improvements'):
            doc.add_heading('🔧 待改进', level=3)
            for imp in script_review['improvements']:
                moment = imp.get('moment', '')
                problem = imp.get('problem', '')
                suggestion = imp.get('suggestion', '')
                if moment:
                    p = doc.add_paragraph()
                    run = p.add_run(f"【{moment}】")
                    run.font.size = Pt(11)
                    run.font.name = 'Arial'
                    run.bold = True
                p = doc.add_paragraph()
                run = p.add_run(f"问题：{problem}")
                run.font.size = Pt(11)
                run.font.name = 'Arial'
                run.bold = True
                p2 = doc.add_paragraph()
                run2 = p2.add_run(f"建议话术：{suggestion}")
                run2.font.size = Pt(11)
                run2.font.name = 'Arial'
                run2.italic = True
                doc.add_paragraph()
        doc.add_paragraph()

    # 五、改进建议
    advice = sections.get('actionable_advice', {})
    if advice:
        add_section_heading(doc, '五', '改进建议')
        if advice.get('immediate'):
            doc.add_heading('⚡ 立即改进（下次就能用）', level=3)
            imm = advice['immediate']
            if imm.get('action'):
                p = doc.add_paragraph()
                run = p.add_run(f"动作：{imm['action']}")
                run.font.size = Pt(11)
                run.font.name = 'Arial'
            if imm.get('script'):
                p2 = doc.add_paragraph()
                run2 = p2.add_run(f"示范话术：{imm['script']}")
                run2.font.size = Pt(11)
                run2.font.name = 'Arial'
                run2.italic = True

        if advice.get('short_term'):
            doc.add_heading('📈 短期提升（本周重点练）', level=3)
            st = advice['short_term']
            if st.get('focus'):
                add_bullet_paragraph(doc, f"重点：{st['focus']}")
            if st.get('practice'):
                add_bullet_paragraph(doc, f"练习方式：{st['practice']}")
            if st.get('success_criteria'):
                add_bullet_paragraph(doc, f"验收标准：{st['success_criteria']}")

        if advice.get('long_term'):
            doc.add_heading('🎯 长期成长（持续优化）', level=3)
            lt = advice['long_term']
            if lt.get('direction'):
                add_bullet_paragraph(doc, f"方向：{lt['direction']}")
            if lt.get('method'):
                add_bullet_paragraph(doc, f"方法：{lt['method']}")
        doc.add_paragraph()

    # 六、管理洞察
    insight = sections.get('management_insight', {})
    if insight:
        add_section_heading(doc, '六', '管理洞察')
        if insight.get('core_problem'):
            p = doc.add_paragraph()
            run = p.add_run(f"核心问题：{insight['core_problem']}")
            run.font.size = Pt(12)
            run.font.name = 'Arial'
            run.bold = True

        if insight.get('coaching_priority'):
            doc.add_heading('辅导优先级', level=3)
            for i, priority in enumerate(insight['coaching_priority'], 1):
                add_bullet_paragraph(doc, f"{i}. {priority}")

        if insight.get('customer_worth'):
            worth = insight['customer_worth']
            doc.add_heading('客户跟进建议', level=3)
            judgment = worth.get('judgment', '')
            reason = worth.get('reason', '')
            p = doc.add_paragraph()
            run = p.add_run(judgment)
            run.font.size = Pt(11)
            run.font.name = 'Arial'
            run.bold = True
            if reason:
                p2 = doc.add_paragraph()
                run2 = p2.add_run(f"理由：{reason}")
                run2.font.size = Pt(11)
                run2.font.name = 'Arial'

        if insight.get('one_on_one_script'):
            doc.add_heading('一对一辅导话术（给主管参考）', level=3)
            p = doc.add_paragraph()
            run = p.add_run(insight['one_on_one_script'])
            run.font.size = Pt(11)
            run.font.name = 'Arial'
            run.italic = True

        if insight.get('team_lesson'):
            lesson = insight['team_lesson']
            doc.add_heading('团队经验', level=3)
            if lesson.get('positive'):
                add_bullet_paragraph(doc, lesson['positive'], '👍')
            if lesson.get('warning'):
                add_bullet_paragraph(doc, lesson['warning'], '⚠️')
        doc.add_paragraph()


def generate_team_report(doc, data, sections):
    """生成批量团队汇总报告"""

    # 一、团队概览
    overview = sections.get('team_overview', {})
    if overview:
        add_section_heading(doc, '一', '团队概览')
        p = doc.add_paragraph()
        run = p.add_run(f"分析范围：{overview.get('total_recordings', '-')} 通录音，{overview.get('total_sales', '-')} 位销售")
        run.font.size = Pt(11)
        run.font.name = 'Arial'

        headers = ['指标', '数值', '环比变化']
        rows = [
            ['平均综合评分', f"{overview.get('avg_score', '-')}/10", overview.get('score_trend', '-')],
            ['高意向客户', f"{overview.get('high_intent', '-')} 通", overview.get('high_intent_pct', '-')],
            ['成交信号', f"{overview.get('closing_signal', '-')} 通", overview.get('closing_signal_pct', '-')],
            ['需重点跟进', f"{overview.get('follow_up', '-')} 通", overview.get('follow_up_pct', '-')],
            ['销售失误', f"{overview.get('mistakes', '-')} 通", overview.get('mistakes_pct', '-')],
        ]
        add_styled_table(doc, headers, rows)
        doc.add_paragraph()

    # 二、标杆萃取
    benchmark = sections.get('benchmark_mining', {})
    if benchmark:
        add_section_heading(doc, '二', '标杆萃取（>=8分录音）')
        count = benchmark.get('count', 0)
        p = doc.add_paragraph()
        run = p.add_run(f"优秀录音共 {count} 通")
        run.font.size = Pt(11)
        run.font.name = 'Arial'
        run.bold = True

        if benchmark.get('success_patterns'):
            doc.add_heading('成功模式', level=3)
            for pattern in benchmark['success_patterns']:
                add_bullet_paragraph(doc, pattern, '🎯')

        if benchmark.get('replicable_scripts'):
            doc.add_heading('可复制话术', level=3)
            for script in benchmark['replicable_scripts']:
                scene = script.get('scene', '')
                content = script.get('content', '')
                p = doc.add_paragraph()
                run = p.add_run(f"【{scene}】")
                run.font.size = Pt(11)
                run.font.name = 'Arial'
                run.bold = True
                p2 = doc.add_paragraph()
                run2 = p2.add_run(content)
                run2.font.size = Pt(11)
                run2.font.name = 'Arial'
                run2.italic = True

        if benchmark.get('best_practices'):
            doc.add_heading('最佳实践清单', level=3)
            for bp in benchmark['best_practices']:
                add_bullet_paragraph(doc, bp, '✅')
        doc.add_paragraph()

    # 三、短板归因
    gap = sections.get('gap_analysis', {})
    if gap:
        add_section_heading(doc, '三', '短板归因（<=4分或失误录音）')
        count = gap.get('count', 0)
        p = doc.add_paragraph()
        run = p.add_run(f"问题录音共 {count} 通")
        run.font.size = Pt(11)
        run.font.name = 'Arial'
        run.bold = True

        if gap.get('problem_map'):
            doc.add_heading('问题分类地图', level=3)
            headers = ['维度', '问题比例', '具体表现']
            rows = []
            for item in gap['problem_map']:
                rows.append([
                    item.get('dimension', ''),
                    item.get('percentage', ''),
                    item.get('symptom', '')
                ])
            add_styled_table(doc, headers, rows)

        if gap.get('top3_mistakes'):
            doc.add_heading('高频失误 TOP3', level=3)
            for i, mistake in enumerate(gap['top3_mistakes'], 1):
                text = f"{mistake.get('description', '')}（{mistake.get('frequency', '-')} 次）"
                add_bullet_paragraph(doc, text, f'{i}.')

        if gap.get('root_causes'):
            doc.add_heading('问题根因', level=3)
            for rc in gap['root_causes']:
                add_bullet_paragraph(doc, rc, '🔍')
        doc.add_paragraph()

    # 四、异议情报
    objection = sections.get('objection_intelligence', {})
    if objection:
        add_section_heading(doc, '四', '异议情报')
        total = objection.get('total_objections', 0)
        p = doc.add_paragraph()
        run = p.add_run(f"本周共记录 {total} 次客户异议")
        run.font.size = Pt(11)
        run.font.name = 'Arial'
        run.bold = True

        if objection.get('distribution'):
            doc.add_heading('异议分布', level=3)
            headers = ['异议类型', '次数', '占比']
            rows = []
            for item in objection['distribution']:
                rows.append([item.get('type', ''), item.get('count', ''), item.get('percentage', '')])
            add_styled_table(doc, headers, rows)

        if objection.get('trends'):
            doc.add_heading('趋势变化', level=3)
            for trend in objection['trends']:
                add_bullet_paragraph(doc, trend, '📈')

        if objection.get('new_objections'):
            doc.add_heading('新出现的异议', level=3)
            for no in objection['new_objections']:
                add_bullet_paragraph(doc, no, '⚠️')

        if objection.get('solutions'):
            doc.add_heading('解决建议库（高频异议话术方案）', level=3)
            for sol in objection['solutions']:
                obj_type = sol.get('objection_type', '')
                p = doc.add_paragraph()
                run = p.add_run(f"【{obj_type}】")
                run.font.size = Pt(11)
                run.font.name = 'Arial'
                run.bold = True
                for variant in sol.get('variants', []):
                    style = variant.get('style', '')
                    content = variant.get('content', '')
                    p2 = doc.add_paragraph()
                    run2 = p2.add_run(f"{style}：{content}")
                    run2.font.size = Pt(10)
                    run2.font.name = 'Arial'
                    run2.italic = True
        doc.add_paragraph()

    # 五、产品表达审计
    product = sections.get('product_audit', {})
    if product:
        add_section_heading(doc, '五', '产品表达审计')
        if product.get('understanding_gaps'):
            doc.add_heading('客户理解偏差', level=3)
            for gap in product['understanding_gaps']:
                add_bullet_paragraph(doc, gap, '🔍')

        if product.get('value_blindspots'):
            doc.add_heading('价值传递盲点', level=3)
            for bs in product['value_blindspots']:
                add_bullet_paragraph(doc, bs, '👁️')

        if product.get('optimization_suggestions'):
            doc.add_heading('话术优化建议', level=3)
            for sug in product['optimization_suggestions']:
                add_bullet_paragraph(doc, sug, '💡')
        doc.add_paragraph()

    # 六、SOP执行审计
    sop = sections.get('sop_audit', {})
    if sop:
        add_section_heading(doc, '六', 'SOP执行审计')
        if sop.get('compliance_rates'):
            doc.add_heading('SOP执行率', level=3)
            headers = ['环节', '执行率']
            rows = []
            for item in sop['compliance_rates']:
                rows.append([item.get('step', ''), item.get('rate', '')])
            add_styled_table(doc, headers, rows)

        if sop.get('bottlenecks'):
            doc.add_heading('SOP卡点', level=3)
            for bn in sop['bottlenecks']:
                add_bullet_paragraph(doc, bn, '⚠️')

        if sop.get('optimization_suggestions'):
            doc.add_heading('SOP优化建议', level=3)
            for sug in sop['optimization_suggestions']:
                add_bullet_paragraph(doc, sug, '🔧')
        doc.add_paragraph()

    # 七、三层清单
    three_tier = sections.get('three_tier_list', {})
    if three_tier:
        add_section_heading(doc, '七', '三层清单')

        if three_tier.get('excellent_recordings'):
            doc.add_heading('🏆 优秀录音（用于培训/话术库收录）', level=3)
            headers = ['录音', '销售', '客户类型', '得分', '核心亮点']
            rows = []
            for rec in three_tier['excellent_recordings']:
                rows.append([
                    rec.get('id', ''),
                    rec.get('sales', ''),
                    rec.get('customer_type', ''),
                    f"{rec.get('score', '-')}/10",
                    rec.get('highlight', '')
                ])
            add_styled_table(doc, headers, rows)

        if three_tier.get('problem_recordings'):
            doc.add_heading('⚠️ 问题录音（用于复盘/一对一）', level=3)
            headers = ['录音', '销售', '客户类型', '得分', '核心问题', '严重程度']
            rows = []
            for rec in three_tier['problem_recordings']:
                rows.append([
                    rec.get('id', ''),
                    rec.get('sales', ''),
                    rec.get('customer_type', ''),
                    f"{rec.get('score', '-')}/10",
                    rec.get('problem', ''),
                    rec.get('severity', '')
                ])
            add_styled_table(doc, headers, rows)

        if three_tier.get('retrain_list'):
            doc.add_heading('🔴 回炉名单（需要干预）', level=3)
            for person in three_tier['retrain_list']:
                name = person.get('name', '')
                reason = person.get('reason', '')
                plan = person.get('plan', '')
                p = doc.add_paragraph()
                run = p.add_run(f"{name}")
                run.font.size = Pt(11)
                run.font.name = 'Arial'
                run.bold = True
                p2 = doc.add_paragraph()
                run2 = p2.add_run(f"原因：{reason}")
                run2.font.size = Pt(10)
                run2.font.name = 'Arial'
                p3 = doc.add_paragraph()
                run3 = p3.add_run(f"回炉计划：{plan}")
                run3.font.size = Pt(10)
                run3.font.name = 'Arial'
                run3.italic = True
                doc.add_paragraph()
        doc.add_paragraph()

    # 八、培训处方
    training = sections.get('training_prescription', {})
    if training:
        add_section_heading(doc, '八', '培训处方')

        if training.get('mandatory'):
            mandatory = training['mandatory']
            doc.add_heading('🎓 本周必修', level=3)
            p = doc.add_paragraph()
            run = p.add_run(f"主题：{mandatory.get('theme', '')}")
            run.font.size = Pt(12)
            run.font.name = 'Arial'
            run.bold = True
            if mandatory.get('target_audience'):
                add_bullet_paragraph(doc, f"目标人群：{mandatory['target_audience']}")
            if mandatory.get('format'):
                add_bullet_paragraph(doc, f"培训形式：{mandatory['format']}")
            if mandatory.get('benchmark_case'):
                add_bullet_paragraph(doc, f"标杆案例：{mandatory['benchmark_case']}")
            if mandatory.get('negative_case'):
                add_bullet_paragraph(doc, f"反面教材：{mandatory['negative_case']}")
            if mandatory.get('materials'):
                add_bullet_paragraph(doc, f"培训材料：{mandatory['materials']}")
            if mandatory.get('acceptance_criteria'):
                add_bullet_paragraph(doc, f"验收标准：{mandatory['acceptance_criteria']}")

        if training.get('optional'):
            doc.add_heading('📚 本周选修', level=3)
            for opt in training['optional']:
                add_bullet_paragraph(doc, opt, '•')
        doc.add_paragraph()

    # 九、管理动作清单
    actions = sections.get('management_actions', {})
    if actions:
        add_section_heading(doc, '九', '明日管理动作')
        if actions.get('morning'):
            doc.add_heading('早上（8:30-9:00）', level=3)
            for item in actions['morning']:
                add_bullet_paragraph(doc, item, '☀️')
        if actions.get('noon'):
            doc.add_heading('中午（12:30-13:00）', level=3)
            for item in actions['noon']:
                add_bullet_paragraph(doc, item, '🌤️')
        if actions.get('evening'):
            doc.add_heading('晚上（17:30-18:00）', level=3)
            for item in actions['evening']:
                add_bullet_paragraph(doc, item, '🌙')
        doc.add_paragraph()

    # 十、单通详情附录
    recordings = data.get('recordings', [])
    if recordings:
        doc.add_page_break()
        add_section_heading(doc, '十', '单通录音详情附录')
        for rec in recordings:
            rec_sections = rec.get('sections', {})
            rec_sales = rec.get('sales_name', '未知')
            rec_customer = rec.get('customer_type', '未知')
            doc.add_heading(f'{rec_sales} · {rec_customer}', level=2)
            generate_single_report(doc, rec, rec_sections)


def generate_report(data):
    """根据分析结果 JSON 生成 Word 报告"""
    doc = Document()

    # 页面设置
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)

    # 页眉
    header = section.header
    header.is_linked_to_previous = False
    header_para = header.paragraphs[0]
    header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_run = header_para.add_run('电销录音管理教练 · 分析报告')
    header_run.font.size = Pt(9)
    header_run.font.color.rgb = RGBColor(128, 128, 128)
    header_run.font.name = 'Arial'

    # 封面标题
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run('电销录音管理教练报告')
    title_run.bold = True
    title_run.font.size = Pt(22)
    title_run.font.name = 'Arial'
    title_run.font.color.rgb = RGBColor(43, 87, 154)

    # 副标题
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sales_name = data.get('sales_name', '未知')
    date_str = data.get('date', datetime.now().strftime('%y.%m.%d'))
    customer_type = data.get('customer_type', '未知')
    subtitle_run = subtitle.add_run(f'{sales_name} · {customer_type}客户 · {date_str}')
    subtitle_run.font.size = Pt(12)
    subtitle_run.font.name = 'Arial'
    subtitle_run.font.color.rgb = RGBColor(100, 100, 100)

    doc.add_paragraph()
    separator = doc.add_paragraph()
    separator.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sep_run = separator.add_run('━' * 40)
    sep_run.font.color.rgb = RGBColor(43, 87, 154)
    sep_run.font.size = Pt(8)
    doc.add_paragraph()

    # 判断是单通报告还是团队报告
    sections = data.get('sections', {})
    is_team_report = bool(
        sections.get('team_overview') or
        sections.get('benchmark_mining') or
        sections.get('gap_analysis')
    )

    if is_team_report:
        generate_team_report(doc, data, sections)
    else:
        generate_single_report(doc, data, sections)

    # 页脚
    footer = section.footer
    footer.is_linked_to_previous = False
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer_para.add_run('电销录音管理教练 · 数据驱动销售增长 | 录音数据不外传，仅用于团队辅导')
    footer_run.font.size = Pt(8)
    footer_run.font.color.rgb = RGBColor(150, 150, 150)
    footer_run.font.name = 'Arial'

    # 保存文件
    output_path = data.get('output_path', '')
    if not output_path:
        output_path = os.path.expanduser(f'~/Desktop/录音洞察/{sales_name}/')

    os.makedirs(output_path, exist_ok=True)

    filename = f'{date_str}_{sales_name}_{customer_type}_录音洞察.docx'
    full_path = os.path.join(output_path, filename)

    doc.save(full_path)

    return full_path


def main():
    if len(sys.argv) < 2:
        print('用法: python3 generate_report.py <分析结果.json>', file=sys.stderr)
        print('将分析结果 JSON 生成为格式化的 Word 报告。', file=sys.stderr)
        sys.exit(1)

    input_path = sys.argv[1]

    try:
        with open(input_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except FileNotFoundError:
        print(f'错误: 文件不存在 "{input_path}"', file=sys.stderr)
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f'错误: JSON 格式无效 - {e}', file=sys.stderr)
        sys.exit(1)

    output_path = generate_report(data)
    print(json.dumps({
        'success': True,
        'output_path': output_path,
        'sales_name': data.get('sales_name', '未知'),
        'customer_type': data.get('customer_type', '未知'),
    }, ensure_ascii=False))


if __name__ == '__main__':
    main()
