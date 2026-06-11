"""
设计项目进度管理系统 - 路由模块
"""
import io
from datetime import datetime
from functools import wraps
from flask import request, jsonify, send_from_directory, session, send_file

from app import app, db
from app.models import (
    User, Designer, ConstructionPlan, PhasePlan, TechnicalPlan,
    MajorCategory, ProjectName, WorkloadStats, generate_name_pinyin
)
from openpyxl import Workbook, load_workbook


def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return jsonify({'code': 401, 'msg': '未登录'}), 401
        return f(*args, **kwargs)
    return decorated_function


def edit_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return jsonify({'code': 401, 'msg': '未登录'}), 401
        user = db.session.get(User, session['user_id'])
        if user.role not in ['admin', 'engineer']:
            return jsonify({'code': 403, 'msg': '权限不足'}), 403
        return f(*args, **kwargs)
    return decorated_function


@app.route('/')
def index():
    return send_from_directory('.', 'index.html')


@app.route('/static/<path:path>')
def send_static(path):
    return send_from_directory('static', path)


EXCEL_HEADERS = {
    'construction': [
        '项目名称', '项目负责人', '单位工程', '成品套图', '设计人',
        '专业', '专业所长', '预估量化指标', '折合A1图纸张数', '计划完成百分比',
        '计划工作日', '计划开始时间', '计划完成时间', '上月滚动偏差', '完成情况', '备注'
    ],
    'phase': [
        '项目名称', '项目负责人', '设计阶段', '成品套图/文件名称', '设计人',
        '专业', '专业所长', '预估量化指标', '折合A1图纸张数', '说明书页数',
        '计划完成百分比', '计划工作日', '计划开始时间', '计划完成时间',
        '上月滚动偏差', '备注'
    ],
    'technical': [
        '项目名称', '项目负责人', '设计阶段', '成品套图/文件名称', '设计人',
        '专业', '专业所长', '预估量化指标', '折合A1图纸张数', '说明书页数',
        '计划完成百分比', '计划工作日', '计划开始时间', '计划完成时间',
        '上月滚动偏差', '备注'
    ],
    'workload': [
        '月份', '所属单位', '姓名', '专业', '专业所长', '人员角色',
        '具体完成工作内容', '量化指标', '项目名称', '项目负责人',
        '设计阶段', '成品套图/文件名称', '折合A1图纸张数', '说明书页数',
        '套图/文件完成率', '实际工作日', '工作开始时间', '工作结束时间',
        '计划偏差(天)', '备注'
    ],
}


from app.routes.auth import *
from app.routes.designers import *
from app.routes.construction import *
from app.routes.phase import *
from app.routes.technical import *
from app.routes.project_names import *
from app.routes.workload import *
