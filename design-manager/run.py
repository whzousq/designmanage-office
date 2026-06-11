"""
设计项目进度管理系统 - 启动入口
"""
from app import app, db

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    
    app.run(host='0.0.0.0', port=5000, debug=True)
