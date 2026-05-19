from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime

db = SQLAlchemy()

class User(db.Model, UserMixin):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(20), nullable=False) # 'admin', 'teacher', 'student'
    related_id = db.Column(db.Integer, nullable=True) # ID of Teacher or Section etc.
    is_active = db.Column(db.Boolean, default=True)
    password_updated_at = db.Column(db.DateTime, nullable=True)
    is_super_admin = db.Column(db.Boolean, default=False)
    
    # Security & Recovery Fields
    failed_login_attempts = db.Column(db.Integer, default=0)
    locked_until = db.Column(db.DateTime, nullable=True)
    security_question = db.Column(db.String(200), nullable=True)
    security_answer = db.Column(db.String(200), nullable=True)
    recovery_otp = db.Column(db.String(6), nullable=True)
    recovery_otp_expiry = db.Column(db.DateTime, nullable=True)
    recovery_email = db.Column(db.String(120), nullable=True)

class Teacher(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    department = db.Column(db.String(50), nullable=False) # JHS, SHS, Both
    # Grade levels stored as comma-separated string e.g. "7,8,9"
    grade_levels = db.Column(db.String(100), nullable=True) 
    max_hours_per_day = db.Column(db.Integer, default=6)
    stay_window_hours = db.Column(db.Integer, default=9)
    is_master = db.Column(db.Boolean, default=False)
    handle_sec_a = db.Column(db.Boolean, default=False)
    is_hybrid = db.Column(db.Boolean, default=False) # Dynamic flag for SHS extension
    is_active = db.Column(db.Boolean, default=True)
    # Preferred days stored as "MWF", "TTH", or "Mon-Fri"
    preferred_days = db.Column(db.String(20), default="Mon-Fri")
    subjects = db.Column(db.String(255), nullable=True) # comma separated subjects

class Classroom(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), nullable=False)
    room_type = db.Column(db.String(20), nullable=False) # Room, Laboratory
    building = db.Column(db.String(20), nullable=False) # JHS, SHS, Both

class Section(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), unique=True, nullable=False)
    department = db.Column(db.String(20), nullable=False) # JHS, SHS
    grade_level = db.Column(db.String(20), nullable=False)
    track = db.Column(db.String(50), nullable=True) # for SHS
    adviser_id = db.Column(db.Integer, db.ForeignKey('teacher.id'), nullable=True)
    room_id = db.Column(db.Integer, db.ForeignKey('classroom.id'), nullable=True)
    is_section_a = db.Column(db.Boolean, default=False)

    adviser = db.relationship('Teacher', backref=db.backref('sections_advised', lazy=True), foreign_keys=[adviser_id])
    room = db.relationship('Classroom', backref=db.backref('home_sections', lazy=True), foreign_keys=[room_id])

class Subject(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    department = db.Column(db.String(20), nullable=False) # JHS, SHS
    requires_lab = db.Column(db.Boolean, default=False)
    duration_mins = db.Column(db.Integer, nullable=False)
    meetings_per_week = db.Column(db.Integer, default=5)
    grade_level = db.Column(db.String(50), nullable=True) # For JHS (e.g. "7-10") or SHS ("11-12")
    track = db.Column(db.String(50), nullable=True) # TVL, STEM, HUMSS etc.
    is_system = db.Column(db.Boolean, default=False)

class Setting(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    # Storing settings as key-value pairs
    key = db.Column(db.String(50), unique=True, nullable=False)
    value = db.Column(db.String(255), nullable=False)

class ScheduleRun(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    created_at = db.Column(db.DateTime, default=datetime.now)
    duration = db.Column(db.Float, default=0.0)
    conflicts = db.Column(db.Integer, default=0)
    conflict_log = db.Column(db.Text, nullable=True)
    is_active = db.Column(db.Boolean, default=False)
    school_year = db.Column(db.String(50), nullable=True)

class Schedule(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    section_id = db.Column(db.Integer, db.ForeignKey('section.id'), nullable=False)
    subject_id = db.Column(db.Integer, db.ForeignKey('subject.id'), nullable=False)
    teacher_id = db.Column(db.Integer, db.ForeignKey('teacher.id'), nullable=False)
    room_id = db.Column(db.Integer, db.ForeignKey('classroom.id'), nullable=False)
    day_of_week = db.Column(db.String(10), nullable=False) # Monday, Tuesday, etc.
    start_time = db.Column(db.String(10), nullable=False) # HH:MM format
    end_time = db.Column(db.String(10), nullable=False) # HH:MM format
    is_soft_break_override = db.Column(db.Boolean, default=False)

    run_id = db.Column(db.Integer, db.ForeignKey('schedule_run.id'), nullable=True)

    section = db.relationship('Section')
    subject = db.relationship('Subject')
    teacher = db.relationship('Teacher')
    room = db.relationship('Classroom')
    run = db.relationship('ScheduleRun', backref=db.backref('schedules', lazy=True, cascade="all, delete-orphan"))


class ActivityLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    actor_username = db.Column(db.String(50), nullable=False)
    role = db.Column(db.String(20), nullable=False) # 'admin', 'teacher', 'student', 'system'
    action = db.Column(db.String(255), nullable=False)
    module = db.Column(db.String(50), nullable=False) # 'Sections', 'Subjects', 'Users', 'Schedule', 'Settings' etc.
    timestamp = db.Column(db.DateTime, default=datetime.now, nullable=False)

from sqlalchemy.event import listens_for

@listens_for(User, 'before_insert')
def before_insert_user(mapper, connection, target):
    if target.role == 'admin':
        from sqlalchemy import select, func
        stmt = select(func.count()).select_from(target.__table__).where(target.__table__.c.role == 'admin')
        res = connection.execute(stmt).scalar()
        if res == 0:
            target.is_super_admin = True
            metadata = target.metadata
            activity_log_table = metadata.tables.get('activity_log')
            if activity_log_table is not None:
                connection.execute(activity_log_table.insert().values(
                    actor_username='system / setup / initial admin creator',
                    role='system',
                    action='Created Super Admin Account',
                    module='Users',
                    timestamp=datetime.now()
                ))
