import os
from app import app, db, User, ActivityLog
from werkzeug.security import generate_password_hash

with app.app_context():
    print("--- Diagnostic verification of Super Admin and Activity Log ---")
    
    # 1. Clean existing test admin accounts so we can test bootstrap logic
    test_admins = User.query.filter(User.username.like("test_admin%")).all()
    for u in test_admins:
        db.session.delete(u)
    db.session.commit()
    print("Cleaned up any previous test_admin accounts.")
    
    # Check current admin count
    admin_count = User.query.filter_by(role='admin').count()
    print(f"Current admins in database: {admin_count}")
    
    # 2. If no admins exist, let's create a test admin and verify promotion
    # Otherwise we'll simulate the check using our listener directly
    new_username = f"test_admin_{admin_count + 1}"
    new_admin = User(
        username=new_username,
        password=generate_password_hash('test1234'),
        role='admin'
    )
    db.session.add(new_admin)
    db.session.commit()
    print(f"Created new admin: {new_username}")
    
    # Fetch it back
    fetched = User.query.filter_by(username=new_username).first()
    print(f"Admin username: {fetched.username}")
    print(f"Is Super Admin: {fetched.is_super_admin}")
    
    # 3. Create a second admin and verify it is NOT a Super Admin
    second_username = f"test_admin_{admin_count + 2}"
    second_admin = User(
        username=second_username,
        password=generate_password_hash('test1234'),
        role='admin'
    )
    db.session.add(second_admin)
    db.session.commit()
    print(f"Created second admin: {second_username}")
    
    # Fetch second back
    fetched_second = User.query.filter_by(username=second_username).first()
    print(f"Second Admin username: {fetched_second.username}")
    print(f"Is Super Admin: {fetched_second.is_super_admin}")
    
    # 4. Check Activity Logs
    logs = ActivityLog.query.order_by(ActivityLog.id.desc()).limit(5).all()
    print("Latest 5 Activity Logs:")
    for log in logs:
        print(f"[{log.timestamp}] {log.actor_username} ({log.role}): {log.action} in {log.module}")
        
    # Cleanup test accounts
    db.session.delete(fetched)
    db.session.delete(fetched_second)
    db.session.commit()
    print("Cleaned up test admin accounts.")
    print("--- Test Complete ---")
