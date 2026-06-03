from pathlib import Path

DEFAULT_WORKSPACE = Path("universal_rename_workspace")
DEFAULT_AI_MODEL = "dsv4flash"
DEFAULT_AI_BASE_URL = "https://api.deepseek.com/v1"
MAX_STEM_LENGTH = 180

WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
}

REPORT_COLUMNS = [
    "原文件名", "输出文件名", "原扩展名", "新文件名主体", "清洗后的新文件名主体",
    "是否保留扩展名", "是否实际改名", "AI建议名", "AI建议原因", "AI置信度",
    "AI审核结果", "AI审核风险", "AI审核说明", "处理说明", "状态",
]

HISTORY_COLUMNS = [
    "时间戳", "原文件名", "输出文件名", "原扩展名", "新文件名主体",
    "清洗后的新文件名主体", "是否保留扩展名", "是否实际改名", "AI建议名",
    "AI建议原因", "AI置信度", "AI审核结果", "AI审核风险", "AI审核说明",
    "处理说明", "状态", "original_prefix_number", "final_prefix_number",
]

RULE_REPORT_COLUMNS = [
    "原文件名", "输出文件名", "原扩展名", "提取编号", "是否命中规则",
    "规则目标主体", "处理方式", "状态",
]

DEFAULT_RULE_RENAME_MAP = {
    "13": "13_m_up",
    "14": "14_m_dw",
    "15": "15_r_up",
    "16": "16_r_dw",
    "17": "17_l_up",
    "18": "18_l_dw",
    "19": "19_next_area",
    "21": "21_two_min",
    "24": "24_brush_clean",
    "25": "25_bth_start",
    "26": "26_guide_l_u",
    "27": "27_guide_r_u",
    "28": "28_guide_l_d",
    "29": "29_guide_r_d",
    "30": "30_time_not_enough",
    "31": "31_half_achieve",
    "32": "32_30s_achieve",
    "33": "33_guide_m_u",
    "34": "34_guide_m_d",
    "78": "78_wake_up",
    "79": "79_encourage",
    "80": "80_slack_off",
}
