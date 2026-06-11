"""
设计项目进度管理系统 - Flask应用初始化
"""
import os
import sys
from datetime import datetime, timedelta
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_cors import CORS

app = Flask(__name__)

app.config['SECRET_KEY'] = 'design_manager_secret_key_2024'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///design_manager.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=8)

db = SQLAlchemy(app)

CORS(app, supports_credentials=True)

from app import models, routes
