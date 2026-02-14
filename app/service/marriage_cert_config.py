"""
结婚证（Marriage Certificate）翻译配置

包含字段映射、关键词列表、性别/国籍翻译等证件特有配置。
与身份证处理模块(image_processor.py)的 FIELD_TRANSLATIONS 等互不干扰。
"""

# 证件关键词（用于识别证件类文本）
MC_CERTIFICATE_KEYWORDS = [
    '姓名', '性别', '国籍', '出生', '身份证', '登记', '证号', '日期', '备注'
]

# 字段翻译映射（中文字段名 → 英文字段名）
MC_FIELD_TRANSLATIONS = {
    '姓名': 'Name',
    '性别': 'Sex',
    '国籍': 'Nationality',
    '出生日期': 'Date of Birth',
    '身份证件号': 'ID Document No.',
    '登记日期': 'Date of Registration',
    '结婚证字号': 'Certificate No.',
    '结婚证号': 'Certificate No.',
    '备注': 'Remarks',
    '持证人': 'Certificate Holder',
    '登记机关': 'Registration Authority',
}

# 性别翻译
MC_GENDER_TRANSLATIONS = {
    '男': 'Male',
    '女': 'Female',
}

# 国籍翻译
MC_NATIONALITY_TRANSLATIONS = {
    '中国': 'Chinese',
    '美国': 'American',
    '英国': 'British',
    '日本': 'Japanese',
    '韩国': 'Korean',
    '法国': 'French',
    '德国': 'German',
}

# 标签字段集合（这些字段左侧是标签，右侧通常有对应的值）
MC_LABEL_FIELDS = frozenset({
    '登记日期', '姓名', '性别', '国籍', '出生日期', '持证人', '登记机关', '备注',
    '结婚证字号', '证件号', '身份证件号', '身份证号', '护照号', '民族', '出生',
    '住址', '户籍', '血型', '学历', '职业', '联系电话', '电话',
})

# 合并关键词（包含这些关键词的连续文本需要合并翻译）
MC_MERGE_KEYWORDS = [
    '民法典',
    '婚姻法',
    '结婚登记',
    '结婚证',
    '申请结婚',
    '予以登记',
    '发给',
    '符合',
    '规定',
    '中华人民共和国',
    '本法',
]
