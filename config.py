# 西安社保公积金缴费比例配置（2024年最新标准）
SOCIAL_INSURANCE_CONFIG = {
    'SALARY': {
        'code': 'SAL',
        'name': '工资'
    },
    'PENSION': {
        'code': 'PEN',
        'name': '养老保险',
        'rate': 0.16,
        'min_base': 3825,
        'max_base': 20826
    },
    'MEDICAL': {
        'code': 'MED',
        'name': '医疗保险',
        'rate': 0.085,
        'min_base': 3825,
        'max_base': 20826
    },
    'UNEMPLOYMENT': {
        'code': 'UEM',
        'name': '失业保险',
        'rate': 0.005,
        'min_base': 3825,
        'max_base': 20826
    },
    'INJURY': {
        'code': 'INJ',
        'name': '工伤保险',
        'rate': 0.004,
        'min_base': 3825,
        'max_base': 20826
    },
    'HOUSING_FUND': {
        'code': 'HF',
        'name': '住房公积金',
        'rate': 0.12,
        'min_base': 3825,
        'max_base': 20826
    },
    'UNION_FEE': {
        'code': 'UF',
        'name': '工会费',
        'rate': 0.002,
        'min_base': 3825,
        'max_base': 20826
    }
}
