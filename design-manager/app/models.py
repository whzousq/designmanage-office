"""
设计项目进度管理系统 - 数据模型
"""
from datetime import datetime
from app import db
from werkzeug.security import generate_password_hash, check_password_hash

try:
    from pypinyin import pinyin, Style
    USE_PYPINYIN = True
except ImportError:
    USE_PYPINYIN = False


class User(db.Model):
    """用户表"""
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    real_name = db.Column(db.String(80), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='staff')  # admin, engineer, staff
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.now)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def to_dict(self):
        return {
            'id': self.id,
            'username': self.username,
            'realName': self.real_name,
            'role': self.role,
            'isActive': self.is_active,
            'createdAt': self.created_at.strftime('%Y-%m-%d %H:%M:%S') if self.created_at else ''
        }


class Designer(db.Model):
    """设计人员表"""
    __tablename__ = 'designers'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), nullable=False)
    name_pinyin = db.Column(db.String(100))
    department = db.Column(db.String(100))
    major = db.Column(db.String(50))
    title = db.Column(db.String(50))
    phone = db.Column(db.String(20))
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.now)

    def to_dict(self):
        return {
            'id': self.id,
            'name': self.name,
            'namePinyin': self.name_pinyin or '',
            'department': self.department,
            'major': self.major,
            'title': self.title,
            'phone': self.phone,
            'isActive': self.is_active,
            'createdAt': self.created_at.strftime('%Y-%m-%d %H:%M:%S') if self.created_at else ''
        }


def generate_name_pinyin(name):
    """根据姓名生成拼音"""
    if not name or USE_PYPINYIN is False:
        return ''
    
    chinese_chars = [c for c in name if '\u4e00' <= c <= '\u9fff']
    if not chinese_chars:
        return ''
    
    try:
        chinese_text = ''.join(chinese_chars)
        pinyin_list = pinyin(chinese_text, style=Style.NORMAL)
        full_pinyin = ''.join([p[0] if p else '' for p in pinyin_list])
        
        if len(chinese_chars) == 2:
            return full_pinyin.lower()
        else:
            surname = pinyin_list[0][0] if pinyin_list and pinyin_list[0] else ''
            initials = ''.join([p[0][0] if p and p[0] else '' for p in pinyin_list[1:]])
            return (surname + initials).lower()
    except:
        return ''


class ConstructionPlan(db.Model):
    """施工图计划表"""
    __tablename__ = 'construction_plans'
    id = db.Column(db.Integer, primary_key=True)
    project_name = db.Column(db.String(200), nullable=False)
    project_manager = db.Column(db.String(50))
    unit_project = db.Column(db.String(200))
    drawing_content = db.Column(db.String(500))
    designer = db.Column(db.String(50))
    major = db.Column(db.String(50))
    major_director = db.Column(db.String(50))
    estimated_quantitative = db.Column(db.String(100))
    a1_drawing_count = db.Column(db.String(50))
    plan_completion_percent = db.Column(db.String(20))
    plan_work_days = db.Column(db.Integer)
    start_date = db.Column(db.String(20))
    end_date = db.Column(db.String(20))
    monthly_deviation = db.Column(db.String(50))
    completion_status = db.Column(db.String(20))
    remarks = db.Column(db.String(500))
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    def to_dict(self):
        return {
            'id': self.id,
            'projectName': self.project_name,
            'projectManager': self.project_manager,
            'unitProject': self.unit_project,
            'drawingContent': self.drawing_content,
            'designer': self.designer,
            'major': self.major,
            'majorDirector': self.major_director,
            'estimatedQuantitative': self.estimated_quantitative,
            'a1DrawingCount': self.a1_drawing_count,
            'planCompletionPercent': self.plan_completion_percent,
            'planWorkDays': self.plan_work_days,
            'startDate': self.start_date,
            'endDate': self.end_date,
            'monthlyDeviation': self.monthly_deviation,
            'completionStatus': self.completion_status,
            'remarks': self.remarks,
            'updatedAt': self.updated_at.strftime('%Y-%m-%d %H:%M:%S') if self.updated_at else ''
        }


class PhasePlan(db.Model):
    """阶段设计计划表"""
    __tablename__ = 'phase_plans'
    id = db.Column(db.Integer, primary_key=True)
    project_name = db.Column(db.String(200), nullable=False)
    project_manager = db.Column(db.String(50))
    design_phase = db.Column(db.String(50))
    product_drawing_name = db.Column(db.String(200))
    designer = db.Column(db.String(50))
    major = db.Column(db.String(200))
    major_director = db.Column(db.String(50))
    estimated_quantitative = db.Column(db.String(100))
    a1_drawing_count = db.Column(db.String(50))
    manual_page_count = db.Column(db.String(50))
    plan_completion_percent = db.Column(db.String(20))
    plan_work_days = db.Column(db.Integer)
    plan_start_date = db.Column(db.String(20))
    plan_end_date = db.Column(db.String(20))
    monthly_deviation = db.Column(db.String(50))
    remarks = db.Column(db.String(500))
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    def to_dict(self):
        return {
            'id': self.id,
            'projectName': self.project_name,
            'projectManager': self.project_manager,
            'designPhase': self.design_phase,
            'productDrawingName': self.product_drawing_name,
            'designer': self.designer,
            'major': self.major,
            'majorDirector': self.major_director,
            'estimatedQuantitative': self.estimated_quantitative,
            'a1DrawingCount': self.a1_drawing_count,
            'manualPageCount': self.manual_page_count,
            'planCompletionPercent': self.plan_completion_percent,
            'planWorkDays': self.plan_work_days if self.plan_work_days else '',
            'planStartDate': self.plan_start_date,
            'planEndDate': self.plan_end_date,
            'monthlyDeviation': self.monthly_deviation,
            'remarks': self.remarks,
            'updatedAt': self.updated_at.strftime('%Y-%m-%d %H:%M:%S') if self.updated_at else ''
        }


class TechnicalPlan(db.Model):
    """技术要求计划表"""
    __tablename__ = 'technical_plans'
    id = db.Column(db.Integer, primary_key=True)
    project_name = db.Column(db.String(200))
    project_manager = db.Column(db.String(50))
    design_phase = db.Column(db.String(50))
    product_drawing_name = db.Column(db.String(200), nullable=False)
    designer = db.Column(db.String(50))
    major = db.Column(db.String(200))
    major_director = db.Column(db.String(50))
    estimated_quantitative = db.Column(db.String(100))
    a1_drawing_count = db.Column(db.String(50))
    manual_page_count = db.Column(db.String(50))
    plan_completion_percent = db.Column(db.String(20))
    plan_work_days = db.Column(db.Integer)
    plan_start_date = db.Column(db.String(20))
    plan_end_date = db.Column(db.String(20))
    monthly_deviation = db.Column(db.String(50))
    remarks = db.Column(db.String(500))
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    def to_dict(self):
        return {
            'id': self.id,
            'projectName': self.project_name,
            'projectManager': self.project_manager,
            'designPhase': self.design_phase,
            'productDrawingName': self.product_drawing_name,
            'designer': self.designer,
            'major': self.major,
            'majorDirector': self.major_director,
            'estimatedQuantitative': self.estimated_quantitative,
            'a1DrawingCount': self.a1_drawing_count,
            'manualPageCount': self.manual_page_count,
            'planCompletionPercent': self.plan_completion_percent,
            'planWorkDays': self.plan_work_days if self.plan_work_days else '',
            'planStartDate': self.plan_start_date,
            'planEndDate': self.plan_end_date,
            'monthlyDeviation': self.monthly_deviation,
            'remarks': self.remarks,
            'updatedAt': self.updated_at.strftime('%Y-%m-%d %H:%M:%S') if self.updated_at else ''
        }


class MajorCategory(db.Model):
    """专业类别表"""
    __tablename__ = 'major_categories'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), nullable=False, unique=True)
    created_at = db.Column(db.DateTime, default=datetime.now)

    def to_dict(self):
        return {
            'id': self.id,
            'name': self.name,
            'createdAt': self.created_at.strftime('%Y-%m-%d %H:%M:%S') if self.created_at else ''
        }


class ProjectName(db.Model):
    """项目名称字典表"""
    __tablename__ = 'project_names'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False, unique=True)
    created_at = db.Column(db.DateTime, default=datetime.now)

    def to_dict(self):
        return {
            'id': self.id,
            'name': self.name,
            'createdAt': self.created_at.strftime('%Y-%m-%d %H:%M:%S') if self.created_at else ''
        }


class WorkloadStats(db.Model):
    """工作量统计表"""
    __tablename__ = 'workload_stats'
    id = db.Column(db.Integer, primary_key=True)
    month = db.Column(db.String(10), nullable=False)
    department = db.Column(db.String(100))
    name = db.Column(db.String(50), nullable=False)
    major = db.Column(db.String(50))
    major_director = db.Column(db.String(50))
    role = db.Column(db.String(50))
    work_content = db.Column(db.String(500))
    quantitative_index = db.Column(db.String(100))
    project_name = db.Column(db.String(200), nullable=False)
    project_manager = db.Column(db.String(50))
    design_phase = db.Column(db.String(50))
    drawing_name = db.Column(db.String(200))
    a1_drawing_count = db.Column(db.String(50))
    manual_page_count = db.Column(db.String(50))
    completion_rate = db.Column(db.String(20))
    actual_work_days = db.Column(db.Float)
    work_start_date = db.Column(db.String(20))
    work_end_date = db.Column(db.String(20))
    plan_deviation = db.Column(db.Float)
    remarks = db.Column(db.String(500))
    confirmed = db.Column(db.Boolean, default=False)
    confirmed_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    confirmed_at = db.Column(db.DateTime)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    def to_dict(self):
        return {
            'id': self.id,
            'month': self.month,
            'department': self.department,
            'name': self.name,
            'major': self.major,
            'majorDirector': self.major_director,
            'role': self.role,
            'workContent': self.work_content,
            'quantitativeIndex': self.quantitative_index,
            'projectName': self.project_name,
            'projectManager': self.project_manager,
            'designPhase': self.design_phase,
            'drawingName': self.drawing_name,
            'a1DrawingCount': self.a1_drawing_count,
            'manualPageCount': self.manual_page_count,
            'completionRate': self.completion_rate,
            'actualWorkDays': self.actual_work_days if self.actual_work_days else '',
            'workStartDate': self.work_start_date,
            'workEndDate': self.work_end_date,
            'planDeviation': self.plan_deviation if self.plan_deviation else '',
            'remarks': self.remarks,
            'confirmed': self.confirmed,
            'confirmedBy': self.confirmed_by,
            'confirmedAt': self.confirmed_at.strftime('%Y-%m-%d %H:%M:%S') if self.confirmed_at else '',
            'createdBy': self.created_by,
            'updatedAt': self.updated_at.strftime('%Y-%m-%d %H:%M:%S') if self.updated_at else ''
        }
