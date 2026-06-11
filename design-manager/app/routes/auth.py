"""
认证相关路由
"""
from app import app, db
from app.models import User
from flask import request, jsonify, session


@app.route('/api/login', methods=['POST'])
def api_login():
    data = request.get_json()
    if not data or not data.get('username') or not data.get('password'):
        return jsonify({'code': 400, 'msg': '用户名或密码不能为空'}), 400
    
    user = User.query.filter_by(username=data['username']).first()
    if not user or not user.check_password(data['password']):
        return jsonify({'code': 401, 'msg': '用户名或密码错误'}), 401
    
    if not user.is_active:
        return jsonify({'code': 403, 'msg': '账户已禁用'}), 403
    
    session['user_id'] = user.id
    session['username'] = user.username
    session['role'] = user.role
    session.permanent = True
    
    return jsonify({
        'code': 200,
        'msg': '登录成功',
        'data': {
            'userId': user.id,
            'username': user.username,
            'realName': user.real_name,
            'role': user.role
        }
    })


@app.route('/api/logout', methods=['POST'])
def api_logout():
    session.clear()
    return jsonify({'code': 200, 'msg': '退出成功'})


@app.route('/api/user', methods=['GET'])
def api_get_user():
    if 'user_id' not in session:
        return jsonify({'code': 401, 'msg': '未登录'}), 401
    
    user = db.session.get(User, session['user_id'])
    if not user:
        session.clear()
        return jsonify({'code': 401, 'msg': '用户不存在'}), 401
    
    return jsonify({
        'code': 200,
        'data': {
            'userId': user.id,
            'username': user.username,
            'realName': user.real_name,
            'role': user.role
        }
    })
