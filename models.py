# models.py
from flask_sqlalchemy import SQLAlchemy

# Create the SQLAlchemy object without an app instance
db = SQLAlchemy()

class Employee(db.Model):
    __tablename__ = 'employees'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(80), nullable=False)
    age = db.Column(db.Integer)
    salary = db.Column(db.Integer)
    years_in_job = db.Column(db.Integer)
    position = db.Column(db.String(80), nullable=False)
    
    def __repr__(self):
        return f'<Employee {self.name}>'

class User(db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password = db.Column(db.String(80), nullable=False)

class Leave(db.Model):
    __tablename__ = 'leaves'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    leave_type = db.Column(db.String(50), nullable=False)
    start_date = db.Column(db.String(20), nullable=False)
    end_date = db.Column(db.String(20), nullable=False)
    explanation = db.Column(db.String(255))
    
    employee = db.relationship(
        'Employee',
        primaryjoin="Leave.employee_id == Employee.id",
        backref='leaves'
    )