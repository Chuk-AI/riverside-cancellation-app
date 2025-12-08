from flask import (
    Flask,
    render_template,
    render_template_string,
    request,
    redirect,
    url_for,
    session,
    flash,
    jsonify,
    send_file,
)
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import (
    datetime,
    timedelta,
    date,
    time,
    date as date_type,
    time as time_type,
)
from dateutil.relativedelta import relativedelta
from collections import defaultdict
import sqlite3
import os
import re
import json
import csv
import io
from functools import wraps
from dotenv import load_dotenv
import pytz
import zoneinfo

# For Excel Exports

import tempfile
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, NamedStyle, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

# Email System Implementation for Riverside Equestrian
import smtplib
import ssl
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import threading
from functools import wraps


load_dotenv()
# ===================================
# APPLICATION SETUP
# ===================================

app = Flask(__name__)
app.secret_key = "riverside-equestrian-secret-key-change-in-production"
app.config["DATABASE"] = "cancellation_system.db"
app.config["TIMEZONE"] = "America/Toronto"

# Set environment timezone for the application
os.environ["TZ"] = "America/Toronto"


def init_db():
    """Initialize the database with required tables and sample data"""
    conn = sqlite3.connect(app.config["DATABASE"])
    conn.row_factory = sqlite3.Row

    try:
        print("Creating database tables...")

        # Enable foreign key constraints
        conn.execute("PRAGMA foreign_keys = ON")

        # Students table (unchanged)
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS students (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                first_name TEXT NOT NULL,
                last_name TEXT NOT NULL,
                parent_first TEXT,
                parent_last TEXT,
                email TEXT UNIQUE NOT NULL,
                phone TEXT,
                membership_level TEXT NOT NULL DEFAULT 'Bronze',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """
        )
        print("✓ Students table created")

        # Admin users table (unchanged)
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS admin_users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'manager',
                first_name TEXT,
                last_name TEXT,
                active BOOLEAN DEFAULT 1,
                last_login TIMESTAMP,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """
        )
        print("✓ Admin users table created")

        # Membership tiers table (unchanged)
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS membership_tiers (
                level TEXT PRIMARY KEY,
                free_notices INTEGER NOT NULL,
                deadline_hours INTEGER NOT NULL,
                deadline_display TEXT NOT NULL,
                active BOOLEAN DEFAULT 1,
                sort_order INTEGER DEFAULT 1,
                color TEXT DEFAULT '#007bff',
                description TEXT,
                welcome_message TEXT,
                policy_message TEXT,
                allow_sequential BOOLEAN DEFAULT 1,
                allow_rescheduling BOOLEAN DEFAULT 1,
                require_approval BOOLEAN DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """
        )
        print("✓ Membership tiers table created")

        # UPDATED Cancellations table with new columns
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS cancellations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id INTEGER NOT NULL,
                lesson_date DATE NOT NULL,
                lesson_time TIME NOT NULL,
                sequential_lessons TEXT,
                reschedule_requested BOOLEAN DEFAULT 0,
                reschedule_preferences TEXT,
                error_report TEXT,
                cancellation_note TEXT,                    -- NEW COLUMN
                charged BOOLEAN DEFAULT 0,
                excluded BOOLEAN DEFAULT 0,
                approved_by TEXT,
                exclusion_reason TEXT,
                manager_notes TEXT,
                status TEXT DEFAULT 'pending',
                deadline_passed BOOLEAN DEFAULT 0,        -- NEW COLUMN
                is_override BOOLEAN DEFAULT 0,            -- NEW COLUMN
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (student_id) REFERENCES students (id) ON DELETE CASCADE
            )
        """
        )
        print("✓ Cancellations table created")

        # System settings table (unchanged)
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS system_settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL,
                description TEXT,
                category TEXT DEFAULT 'general',
                data_type TEXT DEFAULT 'string',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """
        )
        print("✓ System settings table created")

        # Email templates table (unchanged)
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS email_templates (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                subject TEXT NOT NULL,
                body TEXT NOT NULL,
                type TEXT NOT NULL DEFAULT 'client',
                active BOOLEAN DEFAULT 1,
                auto_send BOOLEAN DEFAULT 1,
                priority TEXT DEFAULT 'normal',
                delay_minutes INTEGER DEFAULT 0,
                include_attachment BOOLEAN DEFAULT 0,
                variables_used TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """
        )
        print("✓ Email templates table created")

        # System logs table (unchanged)
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS system_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                user_type TEXT DEFAULT 'unknown',
                action TEXT NOT NULL,
                details TEXT,
                ip_address TEXT,
                user_agent TEXT,
                session_id TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """
        )
        print("✓ System logs table created")

        # Create indexes for better performance (UPDATED with new columns)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_students_email ON students(email)")
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_students_membership ON students(membership_level)"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_cancellations_student ON cancellations(student_id)"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_cancellations_date ON cancellations(lesson_date)"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_cancellations_created ON cancellations(created_at)"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_cancellations_note ON cancellations(cancellation_note)"  # NEW INDEX
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_cancellations_deadline ON cancellations(deadline_passed)"  # NEW INDEX
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_cancellations_override ON cancellations(is_override)"  # NEW INDEX
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_system_logs_user ON system_logs(user_id)"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_system_logs_action ON system_logs(action)"
        )
        print("✓ Database indexes created")

        # Insert default membership tiers (unchanged)
        membership_tiers = [
            (
                "Bronze",
                1,
                18,
                "6pm previous day",
                1,
                1,
                "#cd7f32",
                "Basic membership with standard cancellation policy",
                "Welcome to Bronze membership!",
                "Please review our standard cancellation policy.",
                1,
                1,
                0,
            ),
            (
                "Silver",
                2,
                18,
                "6pm previous day",
                1,
                2,
                "#c0c0c0",
                "Standard membership with enhanced benefits",
                "Welcome to Silver membership!",
                "You have 2 free cancellations per month.",
                1,
                1,
                0,
            ),
            (
                "Gold",
                4,
                2,
                "2 hours before lesson",
                1,
                3,
                "#ffd700",
                "Premium membership with flexible cancellation policy",
                "Welcome to Gold membership!",
                "Enjoy flexible cancellation up to 2 hours before your lesson.",
                1,
                1,
                0,
            ),
            (
                "Intro Package",
                1,
                18,
                "6pm previous day",
                1,
                4,
                "#17a2b8",
                "Introductory package for new students",
                "Welcome! We hope you enjoy your intro lessons.",
                "As an intro student, you have standard cancellation terms.",
                1,
                1,
                0,
            ),
            (
                "Legacy",
                1,
                18,
                "6pm previous day",
                0,
                5,
                "#6c757d",
                "Legacy membership tier (deprecated)",
                "Thank you for your continued loyalty.",
                "Your legacy membership terms apply.",
                1,
                1,
                0,
            ),
            (
                "Guest",
                1,
                18,
                "6pm previous day",
                1,
                6,
                "#28a745",
                "Guest lesson package",
                "Welcome as our guest!",
                "Guest cancellation policy applies.",
                1,
                1,
                0,
            ),
            (
                "Welcome Package",
                1,
                18,
                "6pm previous day",
                1,
                7,
                "#fd7e14",
                "Welcome package for first-time students",
                "Welcome to Riverside Equestrian!",
                "We are excited to have you join us.",
                1,
                1,
                0,
            ),
        ]

        for tier in membership_tiers:
            conn.execute(
                """
                INSERT OR REPLACE INTO membership_tiers 
                (level, free_notices, deadline_hours, deadline_display, active, sort_order, color, description, welcome_message, policy_message, allow_sequential, allow_rescheduling, require_approval)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                tier,
            )
        print("✓ Default membership tiers inserted")

        # Insert default system settings (unchanged)
        system_settings = [
            (
                "system_name",
                "Riverside Equestrian Cancellation System",
                "System display name",
                "general",
                "string",
            ),
            (
                "company_email",
                "managers@riversideequestrian.ca",
                "Main company email",
                "general",
                "email",
            ),
            (
                "website_url",
                "https://www.riversideequestrian.ca",
                "Company website",
                "general",
                "url",
            ),
            (
                "policy_url",
                "https://www.riversideequestrian.ca/cancellations",
                "Cancellation policy URL",
                "general",
                "url",
            ),
            ("timezone", "America/Toronto", "System timezone", "general", "string"),
            (
                "data_retention_months",
                "12",
                "How long to keep cancellation data",
                "business",
                "integer",
            ),
            (
                "max_sequential_lessons",
                "10",
                "Maximum sequential lessons per notice",
                "business",
                "integer",
            ),
            (
                "illness_documentation_days",
                "14",
                "Days to submit illness documentation",
                "business",
                "integer",
            ),
            (
                "advance_booking_days",
                "7",
                "Days before month start for cancellations",
                "business",
                "integer",
            ),
            (
                "session_timeout_minutes",
                "60",
                "User session timeout",
                "security",
                "integer",
            ),
            (
                "max_login_attempts",
                "5",
                "Maximum failed login attempts",
                "security",
                "integer",
            ),
            (
                "lockout_duration_minutes",
                "15",
                "Account lockout duration",
                "security",
                "integer",
            ),
            (
                "email_from_address",
                "noreply@riversideequestrian.ca",
                "Email from address",
                "email",
                "email",
            ),
            (
                "email_from_name",
                "Riverside Equestrian",
                "Email from name",
                "email",
                "string",
            ),
            (
                "backup_retention_days",
                "30",
                "How long to keep backup files",
                "maintenance",
                "integer",
            ),
        ]

        for setting in system_settings:
            conn.execute(
                """
                INSERT OR REPLACE INTO system_settings (key, value, description, category, data_type)
                VALUES (?, ?, ?, ?, ?)
            """,
                setting,
            )
        print("✓ Default system settings inserted")

        # Insert default email templates (unchanged)
        email_templates = [
            (
                "client_confirmation",
                "Client Confirmation",
                "Cancellation Confirmation - {{client_name}}",
                '<p>Dear {{client_name}},</p><p>This email confirms that we have received your lesson cancellation request.</p><p><strong>Cancellation Details:</strong></p><ul><li>Lesson Date: {{lesson_date}}</li><li>Lesson Time: {{lesson_time}}</li><li>Membership Level: {{membership_tier}}</li><li>Cancellation Status: {{cancellation_status}}</li></ul><p>{{status_message}}</p><p>Current cancellation usage for this month:</p><ul><li>Free cancellations used: {{used_cancellations}} of {{allowed_cancellations}}</li><li>Remaining free cancellations: {{remaining_cancellations}}</li></ul><p>If you have any questions about this cancellation or our policy, please contact us.</p><p>Best regards,<br>The Riverside Equestrian Team</p><hr><p class="small text-muted">For more information about our cancellation policy, visit: <a href="{{policy_url}}">{{policy_url}}</a></p>',
                "client",
                1,
                1,
                "normal",
                0,
                0,
                "client_name,lesson_date,lesson_time,membership_tier,cancellation_status,status_message,used_cancellations,allowed_cancellations,remaining_cancellations,policy_url",
            ),
            (
                "manager_notification",
                "Manager Notification",
                "New Cancellation - {{client_name}} - {{lesson_date}}",
                '<p>A new cancellation has been submitted:</p><p><strong>Client:</strong> {{client_name}} ({{membership_tier}})</p><p><strong>Parent:</strong> {{parent_name}}</p><p><strong>Email:</strong> {{client_email}}</p><p><strong>Phone:</strong> {{client_phone}}</p><p><strong>Lesson:</strong> {{lesson_date}} at {{lesson_time}}</p><p><strong>Status:</strong> {{cancellation_status}}</p><p><strong>Will be charged:</strong> {{will_be_charged}}</p><p><strong>Reason:</strong> {{charge_reason}}</p><p><strong>Sequential lessons:</strong> {{sequential_lessons}}</p><p><strong>Reschedule requested:</strong> {{reschedule_requested}}</p><p><strong>Preferences:</strong> {{reschedule_preferences}}</p><p><strong>Error report:</strong> {{error_report}}</p><p><strong>Submission time:</strong> {{submission_time}}</p><hr><p><a href="{{dashboard_url}}">View in Dashboard</a></p>',
                "manager",
                1,
                1,
                "normal",
                0,
                0,
                "client_name,membership_tier,parent_name,client_email,client_phone,lesson_date,lesson_time,cancellation_status,will_be_charged,charge_reason,sequential_lessons,reschedule_requested,reschedule_preferences,error_report,submission_time,dashboard_url",
            ),
            (
                "cancellation_charged",
                "Cancellation Charged",
                "Cancellation Notice - Charge Applied - {{client_name}}",
                "<p>Dear {{client_name}},</p><p>Your cancellation has been processed, and a charge has been applied to your account.</p><p><strong>Lesson Date:</strong> {{lesson_date}}</p><p><strong>Lesson Time:</strong> {{lesson_time}}</p><p><strong>Reason for charge:</strong> {{charge_reason}}</p><p><strong>Charge amount:</strong> As per your membership agreement</p><p>For questions about this charge, please contact us at {{contact_email}}.</p><p>Best regards,<br>The Riverside Equestrian Team</p>",
                "client",
                1,
                1,
                "normal",
                0,
                0,
                "client_name,lesson_date,lesson_time,charge_reason,contact_email",
            ),
            (
                "free_cancellation",
                "Free Cancellation",
                "Free Cancellation Confirmed - {{client_name}}",
                "<p>Dear {{client_name}},</p><p>Your cancellation has been processed at no charge.</p><p><strong>Lesson Date:</strong> {{lesson_date}}</p><p><strong>Lesson Time:</strong> {{lesson_time}}</p><p><strong>Remaining free cancellations this month:</strong> {{remaining_cancellations}}</p><p>Thank you for giving us advance notice.</p><p>Best regards,<br>The Riverside Equestrian Team</p>",
                "client",
                1,
                1,
                "normal",
                0,
                0,
                "client_name,lesson_date,lesson_time,remaining_cancellations",
            ),
        ]

        for template in email_templates:
            conn.execute(
                """
                INSERT OR REPLACE INTO email_templates 
                (id, name, subject, body, type, active, auto_send, priority, delay_minutes, include_attachment, variables_used)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                template,
            )
        print("✓ Default email templates inserted")

        # Insert sample admin users (unchanged)
        from werkzeug.security import generate_password_hash

        admin_password_hash = generate_password_hash("admin123")

        admin_users = [
            (
                "admin@riversideequestrian.ca",
                admin_password_hash,
                "senior_manager",
                "Senior",
                "Admin",
                1,
            ),
            (
                "manager@riversideequestrian.ca",
                admin_password_hash,
                "manager",
                "Manager",
                "User",
                1,
            ),
        ]

        for user in admin_users:
            conn.execute(
                """
                INSERT OR REPLACE INTO admin_users (email, password_hash, role, first_name, last_name, active)
                VALUES (?, ?, ?, ?, ?, ?)
            """,
                user,
            )
        print("✓ Sample admin users inserted")

        # Insert sample students (unchanged)
        sample_students = [
            (
                "Chloe",
                "Chow",
                "Josephine",
                "Tang",
                "chloechow2016@gmail.com",
                "604-505-3333",
                "Silver",
            ),
            (
                "Emma",
                "Johnson",
                "Sarah",
                "Johnson",
                "emma.johnson@email.com",
                "604-123-4567",
                "Gold",
            ),
            (
                "Alex",
                "Smith",
                "Mike",
                "Smith",
                "alex.smith@email.com",
                "604-987-6543",
                "Bronze",
            ),
            (
                "Sofia",
                "Garcia",
                "Maria",
                "Garcia",
                "sofia.garcia@email.com",
                "604-555-0123",
                "Silver",
            ),
            (
                "Liam",
                "Brown",
                "Jennifer",
                "Brown",
                "liam.brown@email.com",
                "604-111-2222",
                "Bronze",
            ),
            (
                "Olivia",
                "Davis",
                "Robert",
                "Davis",
                "olivia.davis@email.com",
                "604-333-4444",
                "Gold",
            ),
            (
                "Noah",
                "Wilson",
                "Lisa",
                "Wilson",
                "noah.wilson@email.com",
                "604-555-6666",
                "Silver",
            ),
            (
                "Ava",
                "Martinez",
                "Carlos",
                "Martinez",
                "ava.martinez@email.com",
                "604-777-8888",
                "Bronze",
            ),
            (
                "Ethan",
                "Anderson",
                "Nicole",
                "Anderson",
                "ethan.anderson@email.com",
                "604-999-0000",
                "Intro Package",
            ),
            (
                "Isabella",
                "Taylor",
                "David",
                "Taylor",
                "isabella.taylor@email.com",
                "604-222-3333",
                "Welcome Package",
            ),
            (
                "Mason",
                "Moore",
                "Amanda",
                "Moore",
                "mason.moore@email.com",
                "604-444-5555",
                "Bronze",
            ),
            (
                "Sophia",
                "Jackson",
                "Kevin",
                "Jackson",
                "sophia.jackson@email.com",
                "604-666-7777",
                "Silver",
            ),
            (
                "Lucas",
                "White",
                "Rachel",
                "White",
                "lucas.white@email.com",
                "604-888-9999",
                "Gold",
            ),
            (
                "Mia",
                "Harris",
                "Jason",
                "Harris",
                "mia.harris@email.com",
                "604-111-0000",
                "Guest",
            ),
            (
                "Aiden",
                "Clark",
                "Melissa",
                "Clark",
                "aiden.clark@email.com",
                "604-222-1111",
                "Legacy",
            ),
        ]

        for student in sample_students:
            conn.execute(
                """
                INSERT OR REPLACE INTO students (first_name, last_name, parent_first, parent_last, email, phone, membership_level)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
                student,
            )
        print("✓ Sample students inserted")

        # UPDATED sample cancellations with new columns
        sample_cancellations = [
            (
                1,
                "2024-08-20",
                "14:00:00",
                "Student was sick",
                0,
                "approved",
                0,
                0,
                "2024-08-15 10:30:00",
            ),
            (
                2,
                "2024-08-22",
                "16:00:00",
                "Need to reschedule",
                1,
                "charged",
                1,
                0,
                "2024-08-16 09:15:00",
            ),
            (
                3,
                "2024-08-25",
                "15:30:00",
                None,
                0,
                "approved",
                0,
                0,
                "2024-08-17 11:45:00",
            ),
            (
                1,
                "2024-08-28",
                "14:00:00",
                "Family emergency",
                0,
                "approved",
                0,
                1,
                "2024-08-18 08:20:00",
            ),
            (
                4,
                "2024-08-30",
                "13:00:00",
                None,
                1,
                "charged",
                1,
                0,
                "2024-08-19 14:10:00",
            ),
            (
                2,
                "2024-09-02",
                "16:00:00",
                "Weather concerns",
                0,
                "pending",
                0,
                0,
                "2024-08-20 15:30:00",
            ),
            (
                5,
                "2024-09-05",
                "11:00:00",
                None,
                0,
                "approved",
                0,
                0,
                "2024-08-21 12:00:00",
            ),
            (
                6,
                "2024-09-08",
                "17:00:00",
                "Schedule conflict",
                1,
                "charged",
                1,
                0,
                "2024-08-22 18:45:00",
            ),
            (
                7,
                "2024-09-10",
                "10:00:00",
                "Car trouble",
                0,
                "pending",
                0,
                0,
                "2024-08-23 09:30:00",
            ),
            (
                8,
                "2024-09-12",
                "18:00:00",
                None,
                0,
                "approved",
                0,
                0,
                "2024-08-24 16:20:00",
            ),
            (
                3,
                "2024-09-15",
                "15:30:00",
                "Work commitment",
                1,
                "charged",
                1,
                0,
                "2024-08-25 12:15:00",
            ),
            (
                9,
                "2024-09-18",
                "09:00:00",
                "Doctor appointment",
                0,
                "approved",
                0,
                1,
                "2024-08-26 14:45:00",
            ),
            (
                10,
                "2024-09-20",
                "13:30:00",
                None,
                0,
                "pending",
                0,
                0,
                "2024-08-27 11:30:00",
            ),
        ]

        for cancellation in sample_cancellations:
            conn.execute(
                """
                INSERT OR REPLACE INTO cancellations (student_id, lesson_date, lesson_time, cancellation_note, charged, status, deadline_passed, is_override, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                cancellation,
            )
        print("✓ Sample cancellations inserted")

        # Insert sample system logs (unchanged)
        sample_logs = [
            (
                1,
                "admin",
                "login",
                "Senior manager login",
                "127.0.0.1",
                "Mozilla/5.0...",
                "session123",
                "2024-08-15 08:00:00",
            ),
            (
                2,
                "admin",
                "login",
                "Manager login",
                "127.0.0.1",
                "Mozilla/5.0...",
                "session124",
                "2024-08-15 08:30:00",
            ),
            (
                1,
                "client",
                "login",
                "Client login: chloechow2016@gmail.com",
                "127.0.0.1",
                "Mozilla/5.0...",
                "session125",
                "2024-08-15 10:00:00",
            ),
            (
                1,
                "client",
                "cancellation_submitted",
                "Lesson: 2024-08-20 14:00:00, Charged: False",
                "127.0.0.1",
                "Mozilla/5.0...",
                "session125",
                "2024-08-15 10:30:00",
            ),
            (
                1,
                "admin",
                "student_updated",
                "Student ID: 5",
                "127.0.0.1",
                "Mozilla/5.0...",
                "session123",
                "2024-08-16 09:00:00",
            ),
            (
                2,
                "client",
                "login",
                "Client login: emma.johnson@email.com",
                "127.0.0.1",
                "Mozilla/5.0...",
                "session126",
                "2024-08-16 09:00:00",
            ),
            (
                2,
                "client",
                "cancellation_submitted",
                "Lesson: 2024-08-22 16:00:00, Charged: True",
                "127.0.0.1",
                "Mozilla/5.0...",
                "session126",
                "2024-08-16 09:15:00",
            ),
        ]

        for log in sample_logs:
            conn.execute(
                """
                INSERT OR REPLACE INTO system_logs (user_id, user_type, action, details, ip_address, user_agent, session_id, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
                log,
            )
        print("✓ Sample system logs inserted")

        # Commit all changes
        conn.commit()
        print("✓ Database initialization completed successfully!")

        # Verify table creation
        cursor = conn.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = [row[0] for row in cursor.fetchall()]
        print(f"✓ Created {len(tables)} tables: {', '.join(tables)}")

        # Verify data insertion
        cursor = conn.execute("SELECT COUNT(*) FROM students")
        student_count = cursor.fetchone()[0]
        print(f"✓ Inserted {student_count} students")

        cursor = conn.execute("SELECT COUNT(*) FROM cancellations")
        cancellation_count = cursor.fetchone()[0]
        print(f"✓ Inserted {cancellation_count} cancellations")

        cursor = conn.execute("SELECT COUNT(*) FROM membership_tiers")
        tier_count = cursor.fetchone()[0]
        print(f"✓ Inserted {tier_count} membership tiers")

        return True

    except Exception as e:
        print(f"❌ Error initializing database: {str(e)}")
        conn.rollback()
        return False
    finally:
        conn.close()


def reset_database():
    """Delete database file and recreate - Safest approach"""
    try:
        db_path = app.config["DATABASE"]

        # Close any existing connections first
        try:
            conn = sqlite3.connect(db_path)
            conn.close()
        except:
            pass

        # Remove the database file if it exists
        if os.path.exists(db_path):
            try:
                os.remove(db_path)
                print(f"Deleted existing database file: {db_path}")
            except OSError as e:
                print(f"Could not delete database file: {e}")
                # If we can't delete the file, try the table-by-table approach
                return reset_database_tables()

        # Create fresh database
        print("Creating fresh database...")
        return init_db()

    except Exception as e:
        print(f"Error in reset_database: {str(e)}")
        import traceback

        traceback.print_exc()
        return False


def reset_database_tables():
    """Fallback method: Drop all tables and recreate database"""
    try:
        conn = sqlite3.connect(app.config["DATABASE"])

        # Get all table names (excluding system tables)
        cursor = conn.execute(
            """SELECT name FROM sqlite_master 
               WHERE type='table' 
               AND name NOT LIKE 'sqlite_%'"""
        )
        tables = [row[0] for row in cursor.fetchall()]

        # Disable foreign key constraints temporarily
        conn.execute("PRAGMA foreign_keys = OFF")

        # Drop all user tables
        for table in tables:
            try:
                conn.execute(f"DROP TABLE IF EXISTS `{table}`")
                print(f"Dropped table: {table}")
            except Exception as e:
                print(f"Warning: Could not drop table {table}: {e}")

        # Drop user-created indexes
        cursor = conn.execute(
            """SELECT name FROM sqlite_master 
               WHERE type='index' 
               AND name NOT LIKE 'sqlite_%'
               AND sql IS NOT NULL"""
        )
        indexes = [row[0] for row in cursor.fetchall()]

        for index in indexes:
            try:
                conn.execute(f"DROP INDEX IF EXISTS `{index}`")
                print(f"Dropped index: {index}")
            except Exception as e:
                print(f"Warning: Could not drop index {index}: {e}")

        # Reset auto-increment sequences
        try:
            conn.execute("DELETE FROM sqlite_sequence")
            print("Reset auto-increment sequences")
        except:
            pass

        # Re-enable foreign key constraints
        conn.execute("PRAGMA foreign_keys = ON")

        conn.commit()
        conn.close()

        # Reinitialize with new schema
        print("Reinitializing database with new schema...")
        return init_db()

    except Exception as e:
        print(f"Error in reset_database_tables: {str(e)}")
        import traceback

        traceback.print_exc()
        return False


def verify_database():
    """Verify database structure and data"""
    try:
        conn = sqlite3.connect(app.config["DATABASE"])
        conn.row_factory = sqlite3.Row

        # Check if all required tables exist
        required_tables = [
            "students",
            "admin_users",
            "membership_tiers",
            "cancellations",
            "system_settings",
            "email_templates",
            "system_logs",
        ]

        cursor = conn.execute("SELECT name FROM sqlite_master WHERE type='table'")
        existing_tables = [row[0] for row in cursor.fetchall()]

        missing_tables = [
            table for table in required_tables if table not in existing_tables
        ]

        if missing_tables:
            print(f"Missing tables: {missing_tables}")
            return False

        # Check if tables have data
        for table in required_tables:
            cursor = conn.execute(f"SELECT COUNT(*) FROM {table}")
            count = cursor.fetchone()[0]
            print(f"Table {table}: {count} records")

        return True

    except Exception as e:
        print(f"Error verifying database: {str(e)}")
        return False
    finally:
        conn.close()


def get_db():
    """Get database connection"""
    conn = sqlite3.connect(app.config["DATABASE"])
    conn.row_factory = sqlite3.Row
    return conn


# ===================================
# UTILITY FUNCTIONS
# ===================================


def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)

    return decorated_function


def admin_required(f):
    """Decorator to require admin access"""

    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_role" not in session or session["user_role"] not in [
            "manager",
            "senior_manager",
        ]:
            flash("Admin access required", "error")
            return redirect(url_for("dashboard"))
        return f(*args, **kwargs)

    return decorated_function


def senior_admin_required(f):
    """Decorator to require senior admin access"""

    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_role" not in session or session["user_role"] != "senior_manager":
            flash("Senior manager access required", "error")
            return redirect(url_for("dashboard"))
        return f(*args, **kwargs)

    return decorated_function


def log_action(action, details=None):
    """Log user action"""
    try:
        conn = get_db()
        conn.execute(
            "INSERT INTO system_logs (user_id, action, details, ip_address) VALUES (?, ?, ?, ?)",
            (session.get("user_id"), action, details, request.remote_addr),
        )
        conn.commit()
        conn.close()
    except:
        pass  # Don't fail if logging fails


def get_app_timezone():
    """Get the application timezone - always use Toronto"""
    return pytz.timezone("America/Toronto")


def toronto_now():
    """Get current datetime in Toronto timezone - simple and direct"""
    toronto_tz = pytz.timezone("America/Toronto")
    utc_now = datetime.utcnow()
    utc_dt = pytz.UTC.localize(utc_now)
    return utc_dt.astimezone(toronto_tz)


def now_in_app_timezone():
    """Get current datetime in Toronto timezone"""
    return toronto_now()


def localize_datetime(dt, from_tz=None):
    """Convert a datetime to Toronto timezone"""
    toronto_tz = pytz.timezone("America/Toronto")

    if dt is None:
        return None

    # If it's a string, parse it first
    if isinstance(dt, str):
        try:
            if "T" in dt:
                dt = datetime.fromisoformat(dt.replace("Z", "+00:00"))
            else:
                dt = datetime.strptime(dt, "%Y-%m-%d %H:%M:%S")
        except (ValueError, TypeError):
            return dt

    # If datetime is naive, assume it's in UTC and convert to Toronto
    if dt.tzinfo is None:
        # Assume naive datetime is in UTC
        dt = pytz.UTC.localize(dt)

    # Convert to Toronto timezone
    return dt.astimezone(toronto_tz)


def format_datetime_for_display(dt):
    """Format datetime for display in Toronto timezone"""
    if dt is None:
        return "Unknown"

    toronto_tz = pytz.timezone("America/Toronto")

    # If it's a string, parse it first
    if isinstance(dt, str):
        try:
            if "T" in dt:
                dt = datetime.fromisoformat(dt.replace("Z", "+00:00"))
            else:
                dt = datetime.strptime(dt, "%Y-%m-%d %H:%M:%S")
                # Assume parsed datetime from database is in Toronto timezone
                dt = toronto_tz.localize(dt)
        except (ValueError, TypeError):
            return str(dt)

    # If it's a naive datetime (from database), assume it's Toronto time
    if isinstance(dt, datetime) and dt.tzinfo is None:
        dt = toronto_tz.localize(dt)

    # Convert to Toronto timezone if needed
    if hasattr(dt, "tzinfo") and dt.tzinfo is not None:
        dt = dt.astimezone(toronto_tz)

    if isinstance(dt, datetime):
        return dt.strftime("%B %d, %Y at %I:%M %p %Z")

    return str(dt)


def format_date_for_display(dt):
    """Format date for display"""
    if dt is None:
        return "Unknown"

    # Convert to app timezone if it's a datetime
    if isinstance(dt, datetime):
        local_dt = localize_datetime(dt)
        return local_dt.strftime("%B %d, %Y")
    elif isinstance(dt, date):
        return dt.strftime("%B %d, %Y")
    elif isinstance(dt, str):
        try:
            if "T" in dt:
                dt_obj = datetime.fromisoformat(dt.replace("Z", "+00:00"))
            else:
                dt_obj = datetime.strptime(dt, "%Y-%m-%d")
            return format_date_for_display(dt_obj)
        except (ValueError, TypeError):
            return str(dt)

    return str(dt)


def get_student_by_email(email):
    """Get student information by email"""
    conn = get_db()
    student = conn.execute(
        "SELECT * FROM students WHERE email = ?", (email,)
    ).fetchone()
    conn.close()
    return student


def get_membership_tier(level):
    """Get membership tier information"""
    conn = get_db()
    tier = conn.execute(
        "SELECT * FROM membership_tiers WHERE level = ?", (level,)
    ).fetchone()
    conn.close()
    return tier


def get_monthly_cancellation_count(student_id, month=None, year=None):
    """Get cancellation count for a student in a specific month"""
    current_time = toronto_now()
    if not month:
        month = current_time.month
    if not year:
        year = current_time.year

    conn = get_db()
    count = conn.execute(
        """
        SELECT COUNT(*) as count FROM cancellations 
        WHERE student_id = ? 
        AND strftime('%m', created_at) = ? 
        AND strftime('%Y', created_at) = ?
        AND excluded = 0
    """,
        (student_id, f"{month:02d}", str(year)),
    ).fetchone()
    conn.close()
    return count["count"] if count else 0


def calculate_cancellation_status(student):
    """Calculate current cancellation status for a student"""
    tier = get_membership_tier(student["membership_level"])
    used = get_monthly_cancellation_count(student["id"])

    return {
        "limit": tier["free_notices"],
        "used": used,
        "remaining": max(0, tier["free_notices"] - used),
    }


def parse_lesson_datetime(lesson_date_str, lesson_time_str):
    """
    Parse lesson date and time strings into a datetime object.
    Handles both HH:MM and HH:MM:SS time formats.
    """
    try:
        # Parse date
        lesson_date = datetime.strptime(str(lesson_date_str), "%Y-%m-%d").date()
    except (ValueError, TypeError):
        # Fallback to today's date if parsing fails
        lesson_date = datetime.now().date()

    try:
        # Parse time - handle both HH:MM and HH:MM:SS formats
        time_str = str(lesson_time_str)
        if len(time_str.split(":")) == 2:  # HH:MM format
            lesson_time = datetime.strptime(time_str, "%H:%M").time()
        elif len(time_str.split(":")) == 3:  # HH:MM:SS format
            lesson_time = datetime.strptime(time_str, "%H:%M:%S").time()
        else:
            # Default to noon if format is unexpected
            lesson_time = time(12, 0)
    except (ValueError, TypeError):
        # Default to noon if parsing fails
        lesson_time = time(12, 0)

    return datetime.combine(lesson_date, lesson_time)


def will_be_charged(student, lesson_datetime):
    """Check if a cancellation will be charged based on submission time vs deadline"""
    tier = get_membership_tier(student["membership_level"])
    status = calculate_cancellation_status(student)

    # Calculate deadline based on lesson date/time and membership tier
    submission_time = toronto_now().replace(
        tzinfo=None
    )  # Convert to naive for comparison

    # For Gold members: 2 hours before lesson
    if tier["level"] == "Gold":
        deadline = lesson_datetime - timedelta(hours=2)
        if submission_time > deadline:
            return True, "Notice submitted less than 2 hours before lesson time"
    else:
        # For all other tiers: 6pm the day before
        lesson_date = lesson_datetime.date()
        previous_day = lesson_date - timedelta(days=1)
        deadline = datetime.combine(previous_day, time(18, 0))  # 6pm previous day

        if submission_time > deadline:
            return True, "Notice submitted after 6pm the previous day"

    # Check monthly limit
    if status["remaining"] <= 0:
        return True, "No more available free cancellation notices this month"

    return False, "This cancellation will be processed as a free cancellation notice"


def get_dashboard_stats():
    """Calculate dashboard statistics with new 5-box metrics and debugging"""
    conn = get_db()

    # Today's date for calculations
    today = datetime.now().date()
    yesterday = today - timedelta(days=1)
    current_month = datetime.now().strftime("%Y-%m")

    try:
        print(f"DEBUG: Today is {today}, current month is {current_month}")

        # Today's cancellations
        today_cancellations_result = conn.execute(
            "SELECT COUNT(*) as count FROM cancellations WHERE DATE(created_at) = ?",
            (today.strftime("%Y-%m-%d"),),
        ).fetchone()
        today_cancellations = (
            today_cancellations_result["count"] if today_cancellations_result else 0
        )
        print(f"DEBUG: Today's cancellations: {today_cancellations}")

        # Yesterday's cancellations for comparison
        yesterday_cancellations_result = conn.execute(
            "SELECT COUNT(*) as count FROM cancellations WHERE DATE(created_at) = ?",
            (yesterday.strftime("%Y-%m-%d"),),
        ).fetchone()
        yesterday_cancellations = (
            yesterday_cancellations_result["count"]
            if yesterday_cancellations_result
            else 0
        )
        print(f"DEBUG: Yesterday's cancellations: {yesterday_cancellations}")

        # Calculate today vs yesterday
        if yesterday_cancellations > 0:
            change = (
                (today_cancellations - yesterday_cancellations)
                / yesterday_cancellations
            ) * 100
            if change > 0:
                today_vs_yesterday = f"+{change:.0f}% from yesterday"
            elif change < 0:
                today_vs_yesterday = f"{change:.0f}% from yesterday"
            else:
                today_vs_yesterday = "Same as yesterday"
        else:
            if today_cancellations > 0:
                today_vs_yesterday = "New activity"
            else:
                today_vs_yesterday = "No activity"

        # This month's cancellations - using date range method
        print(f"DEBUG: Checking month cancellations for {current_month}")

        month_start = f"{current_month}-01"
        next_month = (
            (datetime.now().replace(day=1) + timedelta(days=32))
            .replace(day=1)
            .strftime("%Y-%m-01")
        )

        month_cancellations_result = conn.execute(
            "SELECT COUNT(*) as count FROM cancellations WHERE created_at >= ? AND created_at < ?",
            (month_start, next_month),
        ).fetchone()
        month_cancellations = (
            month_cancellations_result["count"] if month_cancellations_result else 0
        )
        print(f"DEBUG: Month cancellations: {month_cancellations}")

        # This month's detailed stats - UPDATED for 5-box layout
        month_stats = conn.execute(
            """SELECT 
                SUM(CASE WHEN charged = 0 AND excluded = 0 AND status = 'approved' THEN 1 ELSE 0 END) as free,
                SUM(CASE WHEN charged = 1 AND excluded = 0 THEN 1 ELSE 0 END) as charged,
                SUM(CASE WHEN excluded = 1 THEN 1 ELSE 0 END) as excluded,
                SUM(CASE WHEN is_override = 1 THEN 1 ELSE 0 END) as override_count,
                SUM(CASE WHEN deadline_passed = 1 THEN 1 ELSE 0 END) as deadline_passed,
                SUM(CASE WHEN cancellation_note IS NOT NULL AND cancellation_note != '' THEN 1 ELSE 0 END) as with_notes
               FROM cancellations 
               WHERE created_at >= ? AND created_at < ?""",
            (month_start, next_month),
        ).fetchone()

        free_cancellations = int(month_stats["free"] or 0)
        charged_cancellations = int(month_stats["charged"] or 0)
        excluded_cancellations = int(month_stats["excluded"] or 0)
        override_this_month = int(month_stats["override_count"] or 0)
        deadline_passed = int(month_stats["deadline_passed"] or 0)
        with_notes = int(month_stats["with_notes"] or 0)

        print(f"DEBUG: Free: {free_cancellations}, Charged: {charged_cancellations}")
        print(
            f"DEBUG: Excluded: {excluded_cancellations}, Override: {override_this_month}"
        )
        print(f"DEBUG: Deadline passed: {deadline_passed}, With notes: {with_notes}")

        # Active students
        active_students_result = conn.execute(
            "SELECT COUNT(*) as count FROM students"
        ).fetchone()
        active_students = (
            active_students_result["count"] if active_students_result else 0
        )
        print(f"DEBUG: Active students: {active_students}")

        # Sample students for debugging
        students_sample = conn.execute(
            "SELECT first_name, last_name FROM students LIMIT 5"
        ).fetchall()
        student_names = [
            f"{row['first_name']} {row['last_name']}" for row in students_sample
        ]
        print(f"DEBUG: Sample students: {student_names}")

        # New students this month
        try:
            new_students_month_result = conn.execute(
                "SELECT COUNT(*) as count FROM students WHERE created_at >= ? AND created_at < ?",
                (month_start, next_month),
            ).fetchone()
            new_students_month = (
                new_students_month_result["count"] if new_students_month_result else 0
            )
        except Exception as e:
            print(f"DEBUG: Error getting new students: {e}")
            new_students_month = 0

        # Pending reviews - simplified logic
        print("DEBUG: Checking pending reviews...")

        pending_reviews_result = conn.execute(
            """SELECT COUNT(*) as count FROM cancellations 
               WHERE status = 'pending' 
               OR (charged = 1 AND excluded = 0 AND status != 'processed')"""
        ).fetchone()
        pending_reviews = (
            pending_reviews_result["count"] if pending_reviews_result else 0
        )
        print(f"DEBUG: Total pending reviews: {pending_reviews}")

        # Check what statuses we have for debugging
        statuses = conn.execute(
            "SELECT status, COUNT(*) as count FROM cancellations GROUP BY status"
        ).fetchall()
        print(
            f"DEBUG: All statuses: {[(row['status'], row['count']) for row in statuses]}"
        )

        # Active membership tiers
        try:
            active_tiers_result = conn.execute(
                "SELECT COUNT(*) as count FROM membership_tiers WHERE active = 1"
            ).fetchone()
            active_tiers = active_tiers_result["count"] if active_tiers_result else 7
            print(f"DEBUG: Active tiers: {active_tiers}")
        except Exception as e:
            print(f"DEBUG: Error getting active tiers: {e}")
            active_tiers = 7

        # System alerts
        system_alerts = 0
        try:
            old_pending = conn.execute(
                "SELECT COUNT(*) as count FROM cancellations WHERE status = 'pending' AND created_at < date('now', '-3 days')"
            ).fetchone()
            old_pending_count = old_pending["count"] if old_pending else 0
            if old_pending_count > 0:
                system_alerts += 1
        except:
            pass

        # Recent cancellation dates for debugging
        all_dates = conn.execute(
            "SELECT created_at FROM cancellations ORDER BY created_at DESC LIMIT 10"
        ).fetchall()
        print(
            f"DEBUG: Recent cancellation dates: {[row['created_at'] for row in all_dates]}"
        )

        conn.close()

        # Final stats object - UPDATED with new 5-box metrics
        stats = {
            # 5-Box Layout Metrics (Primary)
            "today_cancellations": int(today_cancellations),
            "free_cancellations": int(free_cancellations),
            "charged_cancellations": int(charged_cancellations),
            "excluded_cancellations": int(excluded_cancellations),
            "override_this_month": int(override_this_month),
            # Additional Stats (for other parts of the system)
            "today_vs_yesterday": today_vs_yesterday,
            "month_cancellations": int(month_cancellations),
            "active_students": int(active_students),
            "new_students_month": int(new_students_month),
            "pending_reviews": int(pending_reviews),
            "total_students": int(active_students),
            "monthly_cancellations": int(month_cancellations),
            "active_tiers": int(active_tiers),
            "system_alerts": int(system_alerts),
            # Filter Tab Metrics
            "deadline_passed": int(deadline_passed),
            "with_notes": int(with_notes),
            "total_cancellations": int(month_cancellations),
            # Legacy compatibility
            "today_submissions": int(today_cancellations),
            "urgent_cancellations": int(pending_reviews),
        }

        print(f"DEBUG: Final stats: {stats}")
        return stats

    except Exception as e:
        print(f"DEBUG: Error in get_dashboard_stats: {e}")
        import traceback

        traceback.print_exc()
        conn.close()

        # Return safe defaults if there's an error
        return {
            # 5-Box Layout Metrics
            "today_cancellations": 0,
            "free_cancellations": 0,
            "charged_cancellations": 0,
            "excluded_cancellations": 0,
            "override_this_month": 0,
            # Additional Stats
            "today_vs_yesterday": "Error",
            "month_cancellations": 0,
            "active_students": 0,
            "new_students_month": 0,
            "pending_reviews": 0,
            "total_students": 0,
            "monthly_cancellations": 0,
            "active_tiers": 7,
            "system_alerts": 0,
            # Filter Tab Metrics
            "deadline_passed": 0,
            "with_notes": 0,
            "total_cancellations": 0,
            # Legacy compatibility
            "today_submissions": 0,
            "urgent_cancellations": 0,
        }


# ===================================
# AUTHENTICATION ROUTES
# ===================================


@app.route("/")
def index():
    """Home page - redirect to appropriate dashboard"""
    if "user_id" in session:
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    """Login page for all user types"""
    if request.method == "POST":
        email = request.form["email"].lower().strip()
        password = request.form.get("password", "")

        # Try admin login first
        conn = get_db()
        admin = conn.execute(
            "SELECT * FROM admin_users WHERE email = ?", (email,)
        ).fetchone()

        if admin and check_password_hash(admin["password_hash"], password):
            session["user_id"] = admin["id"]
            session["user_email"] = admin["email"]
            session["user_role"] = admin["role"]
            session["user_name"] = admin["email"]
            conn.close()
            log_action("login", f"Admin login: {admin['role']}")
            flash(f'Welcome back, {admin["role"].title()}!', "success")
            # Redirect managers to cancellations page
            return redirect(url_for("manager_cancellations"))

        # Try student login (email only, no password required)
        student = conn.execute(
            "SELECT * FROM students WHERE email = ?", (email,)
        ).fetchone()
        conn.close()

        if student:
            session["user_id"] = student["id"]
            session["user_email"] = student["email"]
            session["user_role"] = "client"
            session["user_name"] = f"{student['first_name']} {student['last_name']}"
            log_action("login", f"Client login: {student['email']}")
            flash(f'Welcome back, {student["first_name"]}!', "success")
            # Redirect clients to cancel lesson page
            return redirect(url_for("client_cancel"))

        log_action("login_failed", f"Failed login attempt: {email}")
        flash("Invalid email or password", "error")

    return render_template("login.html")


@app.route("/logout")
def logout():
    """Logout user"""
    user_name = session.get("user_name", "User")
    log_action("logout", f"User logged out: {user_name}")
    session.clear()
    flash(f"Goodbye, {user_name}!", "info")
    return redirect(url_for("login"))


# ===================================
# DASHBOARD ROUTES
# ===================================


@app.route("/dashboard")
@login_required
def dashboard():
    """Main dashboard - route to appropriate user dashboard"""
    if session["user_role"] == "client":
        return redirect(url_for("client_dashboard"))
    elif session["user_role"] == "manager":
        return redirect(url_for("manager_dashboard"))
    elif session["user_role"] == "senior_manager":
        return redirect(url_for("senior_dashboard"))
    else:
        return redirect(url_for("login"))


# ===================================
# CLIENT ROUTES
# ===================================


@app.route("/client/dashboard")
@login_required
def client_dashboard():
    """Client dashboard"""
    if session["user_role"] != "client":
        return redirect(url_for("dashboard"))

    conn = get_db()
    client = conn.execute(
        "SELECT * FROM students WHERE id = ?", (session["user_id"],)
    ).fetchone()

    # Get recent cancellations
    recent_cancellations_raw = conn.execute(
        """
        SELECT * FROM cancellations 
        WHERE student_id = ? 
        ORDER BY created_at DESC 
        LIMIT 5
    """,
        (session["user_id"],),
    ).fetchall()
    conn.close()

    if not client:
        flash("Student record not found", "error")
        return redirect(url_for("logout"))

    # Convert string dates to datetime objects for template
    recent_cancellations = []
    for cancellation in recent_cancellations_raw:
        cancellation_dict = dict(cancellation)
        # Convert string dates to datetime objects
        try:
            cancellation_dict["lesson_date"] = datetime.strptime(
                cancellation["lesson_date"], "%Y-%m-%d"
            ).date()

            # Handle both HH:MM and HH:MM:SS time formats
            time_str = str(cancellation["lesson_time"])
            if len(time_str.split(":")) == 2:  # HH:MM format
                cancellation_dict["lesson_time"] = datetime.strptime(
                    time_str, "%H:%M"
                ).time()
            else:  # HH:MM:SS format
                cancellation_dict["lesson_time"] = datetime.strptime(
                    time_str, "%H:%M:%S"
                ).time()

            cancellation_dict["created_at"] = datetime.strptime(
                cancellation["created_at"], "%Y-%m-%d %H:%M:%S"
            )
            # Mark this as Toronto time for display functions
            toronto_tz = pytz.timezone("America/Toronto")
            cancellation_dict["created_at"] = toronto_tz.localize(
                cancellation_dict["created_at"]
            )
        except (ValueError, TypeError):
            # Skip problematic records
            continue
        recent_cancellations.append(cancellation_dict)

    # Calculate status
    cancellation_status = calculate_cancellation_status(client)
    cancellation_policy = get_membership_tier(client["membership_level"])

    return render_template(
        "client_dashboard.html",
        client=client,
        cancellation_status=cancellation_status,
        cancellation_policy=cancellation_policy,
        current_month=datetime.now().strftime("%B %Y"),
        recent_cancellations=recent_cancellations,
    )


@app.route("/client/cancel", methods=["GET", "POST"])
@login_required
def client_cancel():
    """Client cancellation form - UPDATED with cancellation note support"""
    if session["user_role"] != "client":
        return redirect(url_for("dashboard"))

    conn = get_db()
    client = conn.execute(
        "SELECT * FROM students WHERE id = ?", (session["user_id"],)
    ).fetchone()
    conn.close()

    if not client:
        flash("Student record not found", "error")
        return redirect(url_for("logout"))

    if request.method == "POST":
        lesson_date = request.form["lesson_date"]
        lesson_time = request.form["lesson_time"]
        sequential_dates = request.form.getlist("sequential_dates[]")
        sequential_times = request.form.getlist("sequential_times[]")
        wants_reschedule = "wants_reschedule" in request.form
        reschedule_preferences = request.form.get("reschedule_preferences", "")
        error_report = request.form.get("error_report", "")
        cancellation_note = request.form.get("cancellation_note", "")  # NEW FIELD

        # Validate main lesson
        try:
            lesson_datetime = datetime.strptime(
                f"{lesson_date} {lesson_time}", "%Y-%m-%d %H:%M"
            )
        except ValueError:
            flash("Invalid date or time format", "error")
            return redirect(url_for("client_cancel"))

        current_time = toronto_now().replace(
            tzinfo=None
        )  # Convert to naive for comparison

        if lesson_datetime <= current_time:
            flash("Cannot cancel lessons that have already occurred", "error")
            return redirect(url_for("client_cancel"))

        # Check if will be charged
        will_charge, charge_reason = will_be_charged(client, lesson_datetime)

        # Determine deadline status for new database fields
        tier = get_membership_tier(client["membership_level"])
        deadline_hours = tier["deadline_hours"] if tier else 18
        hours_until_lesson = (lesson_datetime - current_time).total_seconds() / 3600
        deadline_passed = hours_until_lesson < deadline_hours

        # Prepare sequential lessons data
        sequential_lessons = []
        if sequential_dates and sequential_times:
            for i, (seq_date, seq_time) in enumerate(
                zip(sequential_dates, sequential_times)
            ):
                if seq_date and seq_time:
                    sequential_lessons.append({"date": seq_date, "time": seq_time})

        sequential_lessons_json = (
            str(sequential_lessons) if sequential_lessons else None
        )

        # Insert cancellation - UPDATED with new fields
        conn = get_db()
        cursor = conn.execute(
            """
            INSERT INTO cancellations
            (student_id, lesson_date, lesson_time, sequential_lessons,
             reschedule_requested, reschedule_preferences, error_report, 
             cancellation_note, charged, deadline_passed, is_override, 
             status, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
            (
                session["user_id"],
                lesson_date,
                lesson_time,
                sequential_lessons_json,
                wants_reschedule,
                reschedule_preferences,
                error_report,
                cancellation_note,  # NEW
                will_charge,
                deadline_passed,  # NEW
                False,  # is_override starts as False
                "pending",  # Default status
                toronto_now().strftime("%Y-%m-%d %H:%M:%S"),  # Use Toronto time
                toronto_now().strftime("%Y-%m-%d %H:%M:%S"),  # Use Toronto time
            ),
        )
        cancellation_id = cursor.lastrowid
        conn.commit()
        conn.close()

        # Prepare cancellation data for email - UPDATED with new fields
        cancellation_data = {
            "id": cancellation_id,
            "lesson_date": lesson_date,
            "lesson_time": lesson_time,
            "sequential_lessons": sequential_lessons_json,
            "reschedule_requested": wants_reschedule,
            "reschedule_preferences": reschedule_preferences,
            "error_report": error_report,
            "cancellation_note": cancellation_note,  # NEW
            "charged": will_charge,
            "deadline_passed": deadline_passed,  # NEW
            "manager_notes": charge_reason,
        }

        # Send confirmation email
        try:
            email_result = send_cancellation_confirmation(client, cancellation_data)

            if email_result["success"]:
                log_action(
                    "email_sent",
                    f"Confirmation email sent for cancellation {cancellation_id}",
                )
            else:
                log_action(
                    "email_failed",
                    f"Failed to send confirmation email: {email_result['message']}",
                )
                print(f"Email sending failed: {email_result['message']}")
        except Exception as e:
            log_action(
                "email_error",
                f"Email system error for cancellation {cancellation_id}: {str(e)}",
            )
            print(f"Email system error: {str(e)}")

        # Send notification email to managers
        try:
            manager_result = send_manager_notification(client, cancellation_data)

            if manager_result["success"]:
                log_action(
                    "email_sent",
                    f"Manager notification sent for cancellation {cancellation_id}: {manager_result['message']}",
                )
                print(f"✅ Manager notification sent: {manager_result['message']}")
            else:
                log_action(
                    "email_failed",
                    f"Failed to send manager notification: {manager_result['message']}",
                )
                print(f"❌ Manager email failed: {manager_result['message']}")
        except Exception as e:
            log_action(
                "email_error",
                f"Manager email system error for cancellation {cancellation_id}: {str(e)}",
            )
            print(f"❌ Manager email error: {str(e)}")

        log_action(
            "cancellation_submitted",
            f"Lesson: {lesson_date} {lesson_time}, Charged: {will_charge}, Note: {'Yes' if cancellation_note else 'No'}",
        )

        # Flash appropriate message - ENHANCED with note acknowledgment
        if will_charge:
            if cancellation_note:
                flash(f"Cancellation submitted with note. {charge_reason}", "warning")
            else:
                flash(f"Cancellation submitted. {charge_reason}", "warning")
        else:
            if cancellation_note:
                flash(
                    "Cancellation with note submitted successfully! This was processed as a free cancellation.",
                    "success",
                )
            else:
                flash(
                    "Cancellation submitted successfully! This was processed as a free cancellation.",
                    "success",
                )

        return redirect(url_for("client_dashboard"))

    # GET request - show form
    cancellation_status = calculate_cancellation_status(client)
    cancellation_policy = get_membership_tier(client["membership_level"])

    return render_template(
        "client_cancel.html",
        client=client,
        cancellation_status=cancellation_status,
        cancellation_policy=cancellation_policy,
        min_date=date.today().isoformat(),
    )


# Backend Route Fix
@app.route("/client/history")
@login_required
def client_history():
    """Client cancellation history"""
    if session["user_role"] != "client":
        return redirect(url_for("dashboard"))
    conn = get_db()
    client = conn.execute(
        "SELECT * FROM students WHERE id = ?", (session["user_id"],)
    ).fetchone()
    # Get all cancellations for this client
    cancellations = conn.execute(
        """
        SELECT * FROM cancellations
        WHERE student_id = ?
        ORDER BY created_at DESC
    """,
        (session["user_id"],),
    ).fetchall()
    conn.close()

    # Group by month
    cancellations_by_month = defaultdict(list)
    for cancellation in cancellations:
        created_date = datetime.strptime(
            cancellation["created_at"], "%Y-%m-%d %H:%M:%S"
        )
        month_key = created_date.strftime("%Y-%m")

        # Parse sequential lessons if any
        sequential_lessons = []
        if cancellation["sequential_lessons"]:
            try:
                sequential_lessons = eval(cancellation["sequential_lessons"])
                # Convert string dates to datetime objects for template
                for lesson in sequential_lessons:
                    if isinstance(lesson.get("date"), str):
                        lesson["date"] = datetime.strptime(
                            lesson["date"], "%Y-%m-%d"
                        ).date()
                    if isinstance(lesson.get("time"), str):
                        try:
                            lesson["time"] = datetime.strptime(
                                lesson["time"], "%H:%M:%S"
                            ).time()
                        except ValueError:
                            lesson["time"] = datetime.strptime(
                                lesson["time"], "%H:%M"
                            ).time()
            except:
                sequential_lessons = []

        cancellation_dict = dict(cancellation)
        cancellation_dict["sequential_lessons"] = sequential_lessons
        cancellation_dict["created_at"] = created_date
        cancellation_dict["lesson_date"] = datetime.strptime(
            cancellation["lesson_date"], "%Y-%m-%d"
        ).date()

        # Fix: Handle both HH:MM and HH:MM:SS time formats
        lesson_time_str = cancellation["lesson_time"]
        try:
            # Try parsing with seconds first
            cancellation_dict["lesson_time"] = datetime.strptime(
                lesson_time_str, "%H:%M:%S"
            ).time()
        except ValueError:
            # If that fails, try parsing without seconds
            try:
                cancellation_dict["lesson_time"] = datetime.strptime(
                    lesson_time_str, "%H:%M"
                ).time()
            except ValueError:
                # If both fail, set a default time
                cancellation_dict["lesson_time"] = datetime.strptime(
                    "00:00", "%H:%M"
                ).time()

        cancellations_by_month[month_key].append(cancellation_dict)

    # Format for template
    formatted_months = []
    for month_key in sorted(cancellations_by_month.keys(), reverse=True):
        month_date = datetime.strptime(month_key, "%Y-%m")
        formatted_months.append(
            {
                "month": month_key,
                "month_display": month_date.strftime("%B %Y"),
                "count": len(cancellations_by_month[month_key]),
                "cancellations": cancellations_by_month[month_key],
            }
        )

    # Current month stats
    current_month_stats = calculate_cancellation_status(client)

    return render_template(
        "client_history.html",
        client=client,
        cancellations=cancellations,
        cancellations_by_month=formatted_months,
        current_month=datetime.now().strftime("%B %Y"),
        current_month_stats=current_month_stats,
        has_more_history=False,
    )


# ===================================
# MANAGER ROUTES
# ===================================


@app.route("/manager/dashboard")
@login_required
@admin_required
def manager_dashboard():
    """Manager dashboard"""
    # Add current_time for template
    current_time = toronto_now().replace(tzinfo=None)  # Convert to naive for template

    # Get properly calculated stats
    stats = get_dashboard_stats()
    conn = get_db()

    # Recent cancellations with student names
    recent_cancellations_raw = conn.execute(
        """
        SELECT c.*, s.first_name, s.last_name, s.membership_level
        FROM cancellations c
        JOIN students s ON c.student_id = s.id
        ORDER BY c.created_at DESC
        LIMIT 10
    """
    ).fetchall()

    # Convert to proper format for template
    recent_cancellations = []
    for cancellation in recent_cancellations_raw:
        cancellation_dict = dict(cancellation)

        # Convert created_at string to datetime
        cancellation_dict["created_at"] = datetime.strptime(
            cancellation["created_at"], "%Y-%m-%d %H:%M:%S"
        )

        # Convert lesson_date string to date object
        cancellation_dict["lesson_date"] = datetime.strptime(
            cancellation["lesson_date"], "%Y-%m-%d"
        ).date()

        # Convert lesson_time string to time object (handle both HH:MM and HH:MM:SS)
        lesson_time_str = cancellation["lesson_time"]
        try:
            cancellation_dict["lesson_time"] = datetime.strptime(
                lesson_time_str, "%H:%M:%S"
            ).time()
        except ValueError:
            try:
                cancellation_dict["lesson_time"] = datetime.strptime(
                    lesson_time_str, "%H:%M"
                ).time()
            except ValueError:
                # Default fallback
                cancellation_dict["lesson_time"] = datetime.strptime(
                    "00:00", "%H:%M"
                ).time()

        # Create student_name for template
        cancellation_dict["student_name"] = (
            f"{cancellation['first_name']} {cancellation['last_name']}"
        )

        recent_cancellations.append(cancellation_dict)

    # Pending actions (charged cancellations, exclusions needed, etc.)
    pending_actions_raw = conn.execute(
        """
        SELECT c.*, s.first_name, s.last_name, s.membership_level
        FROM cancellations c
        JOIN students s ON c.student_id = s.id
        WHERE c.charged = 1 AND c.excluded = 0
        ORDER BY c.created_at DESC
        LIMIT 5
    """
    ).fetchall()

    # Convert pending actions to proper format
    pending_actions = []
    for action in pending_actions_raw:
        action_dict = dict(action)
        action_dict["created_at"] = datetime.strptime(
            action["created_at"], "%Y-%m-%d %H:%M:%S"
        )
        action_dict["lesson_date"] = datetime.strptime(
            action["lesson_date"], "%Y-%m-%d"
        ).date()

        lesson_time_str = action["lesson_time"]
        try:
            action_dict["lesson_time"] = datetime.strptime(
                lesson_time_str, "%H:%M:%S"
            ).time()
        except ValueError:
            try:
                action_dict["lesson_time"] = datetime.strptime(
                    lesson_time_str, "%H:%M"
                ).time()
            except ValueError:
                action_dict["lesson_time"] = datetime.strptime("00:00", "%H:%M").time()

        action_dict["student_name"] = f"{action['first_name']} {action['last_name']}"
        pending_actions.append(action_dict)

    # Monthly stats for chart
    monthly_stats_raw = conn.execute(
        """
        SELECT
            strftime('%Y-%m', created_at) as month,
            COUNT(*) as total,
            SUM(CASE WHEN charged = 0 THEN 1 ELSE 0 END) as free,
            SUM(CASE WHEN charged = 1 THEN 1 ELSE 0 END) as charged
        FROM cancellations
        WHERE created_at >= date('now', '-6 months')
        GROUP BY strftime('%Y-%m', created_at)
        ORDER BY month
    """
    ).fetchall()

    # Prepare chart data - ensure we have data
    if monthly_stats_raw:
        monthly_labels = []
        monthly_totals = []
        monthly_charged = []
        monthly_free = []

        for stat in monthly_stats_raw:
            month_date = datetime.strptime(stat["month"], "%Y-%m")
            monthly_labels.append(month_date.strftime("%B %Y"))
            monthly_totals.append(stat["total"] or 0)
            monthly_charged.append(stat["charged"] or 0)
            monthly_free.append(stat["free"] or 0)
    else:
        # Default empty data if no cancellations
        monthly_labels = ["No Data"]
        monthly_totals = [0]
        monthly_charged = [0]
        monthly_free = [0]

    # Create recent activity from actual recent cancellations
    recent_activity = []
    for cancellation in recent_cancellations[:5]:  # Take first 5
        activity_type = (
            "excluded"
            if cancellation.get("excluded")
            else ("charged" if cancellation.get("charged") else "new")
        )
        activity_title = {
            "new": "New Cancellation",
            "charged": "Charged Cancellation",
            "excluded": "Excluded Cancellation",
        }.get(activity_type, "Cancellation")

        # Calculate time ago
        time_diff = datetime.now() - cancellation["created_at"]
        if time_diff.days > 0:
            time_ago = f"{time_diff.days} day{'s' if time_diff.days != 1 else ''} ago"
        elif time_diff.seconds > 3600:
            hours = time_diff.seconds // 3600
            time_ago = f"{hours} hour{'s' if hours != 1 else ''} ago"
        else:
            minutes = max(1, time_diff.seconds // 60)
            time_ago = f"{minutes} minute{'s' if minutes != 1 else ''} ago"

        recent_activity.append(
            {
                "type": activity_type,
                "title": activity_title,
                "description": f"Lesson on {cancellation['lesson_date'].strftime('%m/%d/%Y')}",
                "student_name": cancellation["student_name"],
                "time_ago": time_ago,
            }
        )

    conn.close()

    return render_template(
        "manager_dashboard.html",
        current_time=current_time,
        stats=stats,
        recent_cancellations=recent_cancellations,
        pending_actions=pending_actions,
        monthly_stats=monthly_stats_raw,
        monthly_labels=monthly_labels,
        monthly_totals=monthly_totals,
        monthly_charged=monthly_charged,
        monthly_free=monthly_free,
        recent_activity=recent_activity,
        alerts=[],  # Add any system alerts here
    )


@app.route("/manager/students", methods=["GET", "POST"])
@login_required
@admin_required
def manager_students():
    """Manager students page"""
    if request.method == "POST":
        # Handle POST requests (add student, bulk import, etc.)
        action = request.form.get("action", "add_student")

        if action == "add_student":
            # Add single student
            conn = get_db()
            try:
                conn.execute(
                    """INSERT INTO students 
                       (first_name, last_name, parent_first, parent_last, email, phone, membership_level, created_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        request.form["first_name"],
                        request.form["last_name"],
                        request.form.get("parent_first", ""),
                        request.form.get("parent_last", ""),
                        request.form["email"],
                        request.form.get("phone", ""),
                        request.form["membership_level"],
                        toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                    ),
                )
                conn.commit()
                conn.close()
                flash("Student added successfully!", "success")
            except Exception as e:
                conn.close()
                flash(f"Error adding student: {str(e)}", "error")

        elif action == "bulk_import":
            # Handle bulk import
            flash("Bulk import functionality would be implemented here", "info")

        return redirect(url_for("manager_students"))

    # GET request - display students
    conn = get_db()

    # Get search and filter parameters
    search = request.args.get("search", "")
    membership_filter = request.args.get("membership", "")
    sort_by = request.args.get("sort", "name")
    page = request.args.get("page", 1, type=int)
    per_page = 12  # Students per page

    # Build the query with filters
    where_conditions = []
    params = []

    if search:
        where_conditions.append(
            "(s.first_name LIKE ? OR s.last_name LIKE ? OR s.email LIKE ? OR s.phone LIKE ?)"
        )
        search_param = f"%{search}%"
        params.extend([search_param, search_param, search_param, search_param])

    if membership_filter:
        where_conditions.append("s.membership_level = ?")
        params.append(membership_filter)

    where_clause = " AND ".join(where_conditions)
    if where_clause:
        where_clause = "WHERE " + where_clause

    # Sort order - Updated with new options
    sort_orders = {
        "name": "s.last_name, s.first_name",
        "email": "s.email",
        "membership": "s.membership_level",
        "recent": "s.created_at DESC",
        "recent_activity": "s.created_at DESC",  # For now, same as recent - can be enhanced later
        "status": "s.membership_level DESC",  # For now, sort by membership - can be enhanced later
    }
    order_by = sort_orders.get(sort_by, "s.last_name, s.first_name")

    # Get students with cancellation statistics
    students_query = f"""
        SELECT s.*,
               COUNT(c.id) as total_cancellations,
               SUM(CASE WHEN strftime('%Y-%m', c.created_at) = strftime('%Y-%m', 'now') THEN 1 ELSE 0 END) as cancellations_this_month,
               SUM(CASE WHEN c.charged = 0 AND c.excluded = 0 AND strftime('%Y-%m', c.created_at) = strftime('%Y-%m', 'now') THEN 1 ELSE 0 END) as free_used_this_month,
               MAX(c.created_at) as last_cancellation_date
        FROM students s
        LEFT JOIN cancellations c ON s.id = c.student_id
        {where_clause}
        GROUP BY s.id
        ORDER BY {order_by}
    """

    students_raw = conn.execute(students_query, params).fetchall()

    # Process students data for template
    students = []
    for student_raw in students_raw:
        student = dict(student_raw)

        # Convert created_at to datetime object if it exists and is a string
        if student.get("created_at") and isinstance(student["created_at"], str):
            try:
                student["created_at"] = datetime.strptime(
                    student["created_at"], "%Y-%m-%d %H:%M:%S"
                )
            except ValueError:
                # If parsing fails, set to current time
                student["created_at"] = datetime.now()
        elif not student.get("created_at"):
            student["created_at"] = datetime.now()

        # Calculate statistics
        student["total_cancellations"] = student.get("total_cancellations") or 0
        student["cancellations_this_month"] = (
            student.get("cancellations_this_month") or 0
        )
        student["free_used_this_month"] = student.get("free_used_this_month") or 0

        # Get membership tier limits
        membership_limits = {
            "Bronze": 1,
            "Silver": 2,
            "Gold": 4,
            "Intro Package": 1,
            "Legacy": 1,
            "Guest": 1,
            "Welcome Package": 1,
        }

        limit = membership_limits.get(student["membership_level"], 1)
        student["free_remaining"] = max(0, limit - student["free_used_this_month"])

        students.append(student)

    # Apply custom sorting for specific sort types
    if sort_by == "recent_activity":
        # Sort by most recent cancellation date, then by created date
        students.sort(
            key=lambda s: (
                (
                    datetime.strptime(
                        s.get("last_cancellation_date", "1900-01-01 00:00:00"),
                        "%Y-%m-%d %H:%M:%S",
                    )
                    if s.get("last_cancellation_date")
                    else datetime.min
                ),
                s["created_at"],
            ),
            reverse=True,
        )
    elif sort_by == "status":
        # Sort by free_remaining (descending - most free cancellations first)
        students.sort(key=lambda s: s["free_remaining"], reverse=True)

    # Pagination
    total_students = len(students)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_students = students[start:end]

    # Create pagination object
    class Pagination:
        def __init__(self, page, per_page, total):
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1 if self.has_prev else None
            self.next_num = page + 1 if self.has_next else None

        def iter_pages(
            self, left_edge=2, left_current=2, right_current=3, right_edge=2
        ):
            last = self.pages
            for num in range(1, last + 1):
                if (
                    num <= left_edge
                    or (self.page - left_current - 1 < num < self.page + right_current)
                    or num > last - right_edge
                ):
                    yield num

    pagination = Pagination(page, per_page, total_students)

    # Calculate statistics for the template - Updated with new stats
    current_month = datetime.now().strftime("%Y-%m")

    # Total students
    total_students_count = len(students)

    # Students active this month (had cancellations)
    active_this_month = len([s for s in students if s["cancellations_this_month"] > 0])

    # New students this month
    new_this_month = len(
        [s for s in students if s["created_at"].strftime("%Y-%m") == current_month]
    )

    # Membership breakdown - Updated with all membership types
    gold_members = len([s for s in students if s["membership_level"] == "Gold"])
    silver_members = len([s for s in students if s["membership_level"] == "Silver"])
    bronze_members = len([s for s in students if s["membership_level"] == "Bronze"])
    welcome_members = len(
        [s for s in students if s["membership_level"] == "Welcome Package"]
    )
    guest_members = len([s for s in students if s["membership_level"] == "Guest"])

    stats = {
        "total_students": total_students_count,
        "active_this_month": active_this_month,
        "new_this_month": new_this_month,
        "gold_members": gold_members,
        "silver_members": silver_members,
        "bronze_members": bronze_members,  # NEW
        "welcome_members": welcome_members,  # NEW
        "guest_members": guest_members,  # NEW
    }

    # Get membership tiers for dropdown (use static list if table doesn't exist)
    try:
        membership_tiers_raw = conn.execute(
            "SELECT level FROM membership_tiers WHERE active = 1 ORDER BY sort_order"
        ).fetchall()
        membership_tiers = [tier["level"] for tier in membership_tiers_raw]
    except:
        # Fallback to static list if table doesn't exist
        membership_tiers = [
            "Bronze",
            "Silver",
            "Gold",
            "Intro Package",
            "Legacy",
            "Guest",
            "Welcome Package",
        ]

    conn.close()

    return render_template(
        "manager_students.html",
        students=paginated_students,
        membership_tiers=membership_tiers,
        stats=stats,
        pagination=pagination,
        search=search,
        current_membership_filter=membership_filter,
        current_sort=sort_by,
    )


def process_cancellation_dates(cancellation):
    """Process cancellation dates to ensure consistent formatting for templates"""
    cancellation_dict = (
        dict(cancellation) if hasattr(cancellation, "keys") else cancellation
    )

    # Convert created_at string to datetime
    if isinstance(cancellation_dict.get("created_at"), str):
        try:
            cancellation_dict["created_at"] = datetime.strptime(
                cancellation_dict["created_at"], "%Y-%m-%d %H:%M:%S"
            )
        except (ValueError, TypeError):
            cancellation_dict["created_at"] = datetime.now()

    # Convert lesson_date string to date object
    if isinstance(cancellation_dict.get("lesson_date"), str):
        try:
            cancellation_dict["lesson_date"] = datetime.strptime(
                cancellation_dict["lesson_date"], "%Y-%m-%d"
            ).date()
        except (ValueError, TypeError):
            cancellation_dict["lesson_date"] = datetime.now().date()

    # Convert lesson_time string to time object (handle both HH:MM and HH:MM:SS)
    if isinstance(cancellation_dict.get("lesson_time"), str):
        lesson_time_str = cancellation_dict["lesson_time"]
        try:
            cancellation_dict["lesson_time"] = datetime.strptime(
                lesson_time_str, "%H:%M:%S"
            ).time()
        except ValueError:
            try:
                cancellation_dict["lesson_time"] = datetime.strptime(
                    lesson_time_str, "%H:%M"
                ).time()
            except ValueError:
                cancellation_dict["lesson_time"] = datetime.strptime(
                    "00:00", "%H:%M"
                ).time()

    # Process submitted date and time
    if cancellation_dict.get("created_at"):
        created_at = cancellation_dict["created_at"]
        cancellation_dict["submitted_date"] = created_at.date()
        cancellation_dict["submitted_time"] = created_at.strftime("%I:%M %p")

        # Calculate time ago
        now = datetime.now()
        time_diff = now - created_at
        if time_diff.days > 0:
            cancellation_dict["time_ago"] = (
                f"{time_diff.days} day{'s' if time_diff.days != 1 else ''} ago"
            )
        elif time_diff.seconds > 3600:
            hours = time_diff.seconds // 3600
            cancellation_dict["time_ago"] = (
                f"{hours} hour{'s' if hours != 1 else ''} ago"
            )
        else:
            minutes = max(1, time_diff.seconds // 60)
            cancellation_dict["time_ago"] = (
                f"{minutes} minute{'s' if minutes != 1 else ''} ago"
            )

    # Process sequential lessons
    if cancellation_dict.get("sequential_lessons"):
        try:
            if isinstance(cancellation_dict["sequential_lessons"], str):
                sequential_lessons = eval(cancellation_dict["sequential_lessons"])
            else:
                sequential_lessons = cancellation_dict["sequential_lessons"]

            # Convert string dates/times to objects for template consistency
            for lesson in sequential_lessons:
                if isinstance(lesson.get("date"), str):
                    lesson["date"] = datetime.strptime(
                        lesson["date"], "%Y-%m-%d"
                    ).date()
                if isinstance(lesson.get("time"), str):
                    try:
                        lesson["time"] = datetime.strptime(
                            lesson["time"], "%H:%M:%S"
                        ).time()
                    except ValueError:
                        lesson["time"] = datetime.strptime(
                            lesson["time"], "%H:%M"
                        ).time()

            cancellation_dict["sequential_lessons"] = sequential_lessons
        except:
            cancellation_dict["sequential_lessons"] = []

    # Add computed status properties for analytics
    cancellation_dict["is_free"] = (
        not cancellation_dict.get("charged", False)
        and not cancellation_dict.get("excluded", False)
        and cancellation_dict.get("status") == "approved"
    )

    cancellation_dict["is_charged"] = cancellation_dict.get(
        "charged", False
    ) and not cancellation_dict.get("excluded", False)

    # Ensure boolean fields are properly set
    cancellation_dict["is_override"] = bool(cancellation_dict.get("is_override", False))
    cancellation_dict["deadline_passed"] = bool(
        cancellation_dict.get("deadline_passed", False)
    )
    cancellation_dict["excluded"] = bool(cancellation_dict.get("excluded", False))

    return cancellation_dict


@app.route("/manager/api/cancellation/process", methods=["POST"])
@login_required
@admin_required
def process_cancellation():
    """Process individual cancellation - UPDATED with new approve as policy functionality"""
    data = request.json
    action = data.get("action")  # 'approve_policy', 'force_free', 'force_charge'
    cancellation_id = data.get("cancellation_id")
    reason = data.get("reason", "")

    if not action or not cancellation_id:
        return jsonify({"success": False, "message": "Missing required fields"})

    # Only force actions require a reason
    if action in ["force_free", "force_charge"] and not reason.strip():
        return jsonify(
            {
                "success": False,
                "message": "Override reason is required for force actions",
            }
        )

    conn = get_db()

    try:
        # Get cancellation and student data BEFORE updating
        cancellation_data = conn.execute(
            """
            SELECT c.*, s.first_name, s.last_name, s.email, s.phone, s.membership_level,
                   s.parent_first, s.parent_last
            FROM cancellations c
            JOIN students s ON c.student_id = s.id
            WHERE c.id = ?
            """,
            (cancellation_id,),
        ).fetchone()

        if not cancellation_data:
            return jsonify({"success": False, "message": "Cancellation not found"})

        # Prepare student and cancellation data
        student_dict = {
            "id": cancellation_data["student_id"],
            "first_name": cancellation_data["first_name"],
            "last_name": cancellation_data["last_name"],
            "email": cancellation_data["email"],
            "phone": cancellation_data["phone"],
            "membership_level": cancellation_data["membership_level"],
            "parent_first": cancellation_data["parent_first"],
            "parent_last": cancellation_data["parent_last"],
        }

        # Get lesson datetime for policy check
        lesson_datetime = parse_lesson_datetime(
            cancellation_data["lesson_date"], cancellation_data["lesson_time"]
        )

        # Update database based on action
        if action == "approve_policy":
            # Check policy to determine if should be free or charged
            should_be_charged, charge_reason = will_be_charged(
                student_dict, lesson_datetime
            )

            if should_be_charged:
                # Process as charged according to policy
                conn.execute(
                    """UPDATE cancellations 
                       SET charged = 1, status = 'charged', manager_notes = ?, 
                           approved_by = ?, updated_at = ? 
                       WHERE id = ?""",
                    (
                        f"Processed according to policy: {charge_reason}",
                        session.get("user_email", "Unknown Manager"),
                        toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                        cancellation_id,
                    ),
                )
                updated_cancellation = dict(cancellation_data)
                updated_cancellation.update(
                    {
                        "charged": 1,
                        "status": "charged",
                        "manager_notes": f"Processed according to policy: {charge_reason}",
                        "approved_by": session.get("user_email", "Unknown Manager"),
                    }
                )
                log_message = f"Cancellation {cancellation_id} charged according to policy: {charge_reason}"
            else:
                # Process as free according to policy
                conn.execute(
                    """UPDATE cancellations 
                       SET status = 'approved', charged = 0, manager_notes = ?, 
                           approved_by = ?, updated_at = ? 
                       WHERE id = ?""",
                    (
                        f"Processed according to policy: {charge_reason}",
                        session.get("user_email", "Unknown Manager"),
                        toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                        cancellation_id,
                    ),
                )
                updated_cancellation = dict(cancellation_data)
                updated_cancellation.update(
                    {
                        "charged": 0,
                        "status": "approved",
                        "manager_notes": f"Processed according to policy: {charge_reason}",
                        "approved_by": session.get("user_email", "Unknown Manager"),
                    }
                )
                log_message = f"Cancellation {cancellation_id} approved as free according to policy: {charge_reason}"

        elif action == "force_free":
            # Force as free (override policy)
            conn.execute(
                """UPDATE cancellations 
                   SET status = 'approved', charged = 0, manager_notes = ?, 
                       is_override = 1, approved_by = ?, updated_at = ? 
                   WHERE id = ?""",
                (
                    f"Manager Override (Force Free): {reason}",
                    session.get("user_email", "Unknown Manager"),
                    toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                    cancellation_id,
                ),
            )
            updated_cancellation = dict(cancellation_data)
            updated_cancellation.update(
                {
                    "charged": 0,
                    "status": "approved",
                    "is_override": 1,
                    "manager_notes": f"Manager Override (Force Free): {reason}",
                    "approved_by": session.get("user_email", "Unknown Manager"),
                }
            )
            log_message = (
                f"Cancellation {cancellation_id} forced to free (Override: {reason})"
            )

        elif action == "force_charge":
            # Force as charged (override policy)
            conn.execute(
                """UPDATE cancellations 
                   SET charged = 1, status = 'charged', manager_notes = ?, 
                       is_override = 1, approved_by = ?, updated_at = ? 
                   WHERE id = ?""",
                (
                    f"Manager Override (Force Charge): {reason}",
                    session.get("user_email", "Unknown Manager"),
                    toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                    cancellation_id,
                ),
            )
            updated_cancellation = dict(cancellation_data)
            updated_cancellation.update(
                {
                    "charged": 1,
                    "status": "charged",
                    "is_override": 1,
                    "manager_notes": f"Manager Override (Force Charge): {reason}",
                    "approved_by": session.get("user_email", "Unknown Manager"),
                }
            )
            log_message = (
                f"Cancellation {cancellation_id} forced to charge (Override: {reason})"
            )

        else:
            return jsonify({"success": False, "message": "Invalid action"})

        conn.commit()

        # SEND EMAIL NOTIFICATIONS - Only for force actions (overrides)
        if action in ["force_free", "force_charge"]:
            email_results = send_override_notification_emails(
                student_dict,
                updated_cancellation,
                action,
                reason,
                session.get("user_email", "Unknown Manager"),
            )
        else:
            # For approve_policy, no override emails needed (normal workflow)
            email_results = {
                "client_email": {
                    "success": True,
                    "message": "No override email needed",
                },
                "manager_notification": {
                    "success": True,
                    "message": "No override email needed",
                },
            }

        conn.close()

        # Log the action
        log_action("cancellation_processed", log_message)

        return jsonify(
            {
                "success": True,
                "message": f"Cancellation processed successfully",
                "email_summary": email_results,
            }
        )

    except Exception as e:
        conn.close()
        return jsonify({"success": False, "message": str(e)})

        # Log email results
        if email_results["client_email"].get("success"):
            log_action(
                "override_client_email",
                f"Override confirmation sent to {student_dict['email']}",
            )
        else:
            log_action(
                "override_client_email_failed",
                f"Failed to send override confirmation: {email_results['client_email'].get('message', 'Unknown error')}",
            )

        if email_results["manager_notification"].get("success"):
            log_action(
                "override_manager_email", f"Override notification sent to managers"
            )
        else:
            log_action(
                "override_manager_email_failed",
                f"Failed to send override notification: {email_results['manager_notification'].get('message', 'Unknown error')}",
            )

        # Return success with email status
        return jsonify(
            {
                "success": True,
                "message": f"Cancellation {action}d successfully",
                "email_status": {
                    "client_notified": email_results["client_email"].get(
                        "success", False
                    ),
                    "managers_notified": email_results["manager_notification"].get(
                        "success", False
                    ),
                },
            }
        )

    except Exception as e:
        conn.close()
        return jsonify({"success": False, "message": str(e)})


@app.route("/manager/api/cancellation/batch", methods=["POST"])
@login_required
@admin_required
def batch_process_cancellations():
    """Batch process multiple cancellations - UPDATED with new approve as policy functionality"""
    data = request.json
    action = data.get(
        "action"
    )  # 'approve_policy', 'force_free', 'charge', or 'exclude'
    cancellation_ids = data.get("cancellation_ids", [])
    reason = data.get("reason", "Batch processing")

    if not action or not cancellation_ids:
        return jsonify({"success": False, "message": "Missing required fields"})

    # Force actions require a reason
    if action in ["force_free"] and not reason.strip():
        return jsonify(
            {"success": False, "message": "Reason is required for force free action"}
        )

    conn = get_db()
    email_summary = {
        "client_emails_sent": 0,
        "client_emails_failed": 0,
        "manager_notifications_sent": 0,
    }

    try:
        processed_count = 0

        for cancellation_id in cancellation_ids:
            # Get cancellation and student data
            cancellation_data = conn.execute(
                """
                SELECT c.*, s.first_name, s.last_name, s.email, s.phone, s.membership_level,
                       s.parent_first, s.parent_last
                FROM cancellations c
                JOIN students s ON c.student_id = s.id
                WHERE c.id = ?
                """,
                (cancellation_id,),
            ).fetchone()

            if not cancellation_data:
                continue

            # Prepare data
            student_dict = {
                "id": cancellation_data["student_id"],
                "first_name": cancellation_data["first_name"],
                "last_name": cancellation_data["last_name"],
                "email": cancellation_data["email"],
                "phone": cancellation_data["phone"],
                "membership_level": cancellation_data["membership_level"],
                "parent_first": cancellation_data["parent_first"],
                "parent_last": cancellation_data["parent_last"],
            }

            # Get lesson datetime for policy check
            lesson_datetime = parse_lesson_datetime(
                cancellation_data["lesson_date"], cancellation_data["lesson_time"]
            )

            # Update database based on action
            if action == "approve_policy":
                # Check policy to determine if should be free or charged
                should_be_charged, charge_reason = will_be_charged(
                    student_dict, lesson_datetime
                )

                if should_be_charged:
                    # Process as charged according to policy
                    conn.execute(
                        """UPDATE cancellations 
                           SET charged = 1, status = 'charged', manager_notes = ?, 
                               approved_by = ?, updated_at = ? 
                           WHERE id = ?""",
                        (
                            f"Batch processed according to policy: {charge_reason}",
                            session.get("user_email", "Unknown Manager"),
                            toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                            cancellation_id,
                        ),
                    )
                    updated_cancellation = dict(cancellation_data)
                    updated_cancellation.update(
                        {
                            "charged": 1,
                            "status": "charged",
                            "approved_by": session.get("user_email", "Unknown Manager"),
                        }
                    )
                else:
                    # Process as free according to policy
                    conn.execute(
                        """UPDATE cancellations 
                           SET status = 'approved', charged = 0, manager_notes = ?, 
                               approved_by = ?, updated_at = ? 
                           WHERE id = ?""",
                        (
                            f"Batch processed according to policy: {charge_reason}",
                            session.get("user_email", "Unknown Manager"),
                            toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                            cancellation_id,
                        ),
                    )
                    updated_cancellation = dict(cancellation_data)
                    updated_cancellation.update(
                        {
                            "charged": 0,
                            "status": "approved",
                            "approved_by": session.get("user_email", "Unknown Manager"),
                        }
                    )

            elif action == "force_free":
                conn.execute(
                    """UPDATE cancellations 
                       SET status = 'approved', charged = 0, manager_notes = ?, 
                           is_override = 1, approved_by = ?, updated_at = ? 
                       WHERE id = ?""",
                    (
                        f"Batch force free: {reason}",
                        session.get("user_email", "Unknown Manager"),
                        toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                        cancellation_id,
                    ),
                )
                updated_cancellation = dict(cancellation_data)
                updated_cancellation.update(
                    {
                        "charged": 0,
                        "status": "approved",
                        "is_override": 1,
                        "approved_by": session.get("user_email", "Unknown Manager"),
                    }
                )

            elif action == "charge":
                conn.execute(
                    """UPDATE cancellations 
                       SET charged = 1, status = 'charged', manager_notes = ?, 
                           is_override = 1, approved_by = ?, updated_at = ? 
                       WHERE id = ?""",
                    (
                        f"Batch charge: {reason}",
                        session.get("user_email", "Unknown Manager"),
                        toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                        cancellation_id,
                    ),
                )
                updated_cancellation = dict(cancellation_data)
                updated_cancellation.update(
                    {
                        "charged": 1,
                        "status": "charged",
                        "is_override": 1,
                        "approved_by": session.get("user_email", "Unknown Manager"),
                    }
                )

            elif action == "exclude":
                conn.execute(
                    """UPDATE cancellations 
                       SET excluded = 1, exclusion_reason = ?, approved_by = ?, 
                           is_override = 1, updated_at = ? 
                       WHERE id = ?""",
                    (
                        reason,
                        session["user_email"],
                        toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                        cancellation_id,
                    ),
                )
                updated_cancellation = dict(cancellation_data)
                updated_cancellation.update({"excluded": 1, "is_override": 1})

            # Send email notifications for each processed cancellation
            # Only send override emails for force actions, not for approve_policy
            if action in [
                "force_free",
                "charge",
            ]:  # Don't send emails for approve_policy or exclusions
                email_results = send_override_notification_emails(
                    student_dict,
                    updated_cancellation,
                    action,
                    f"Batch {action}: {reason}",
                    session.get("user_email", "Unknown Manager"),
                )

                # Track email results
                if email_results["client_email"].get("success"):
                    email_summary["client_emails_sent"] += 1
                else:
                    email_summary["client_emails_failed"] += 1

            processed_count += 1

        conn.commit()
        conn.close()

        # Send single manager notification about batch operation
        email_summary["manager_notifications_sent"] = 1  # Simplified for batch

        log_action(
            "batch_processing",
            f"Batch {action}: {processed_count} cancellations ({reason})",
        )

        return jsonify(
            {
                "success": True,
                "processed": processed_count,
                "email_summary": email_summary,
            }
        )

    except Exception as e:
        conn.close()
        return jsonify({"success": False, "message": str(e)})


@app.route("/manager/api/cancellation/process-all-pending", methods=["POST"])
@login_required
@admin_required
def process_all_pending_cancellations():
    """Process all pending cancellations according to policy"""
    try:
        conn = get_db()

        # Get all pending cancellations
        pending_cancellations = conn.execute(
            """
            SELECT c.*, s.first_name, s.last_name, s.email, s.phone, s.membership_level,
                   s.parent_first, s.parent_last
            FROM cancellations c
            JOIN students s ON c.student_id = s.id
            WHERE c.status = 'pending' OR c.status IS NULL
            """
        ).fetchall()

        processed_count = 0

        for cancellation_data in pending_cancellations:
            # Prepare student data
            student_dict = {
                "id": cancellation_data["student_id"],
                "first_name": cancellation_data["first_name"],
                "last_name": cancellation_data["last_name"],
                "email": cancellation_data["email"],
                "phone": cancellation_data["phone"],
                "membership_level": cancellation_data["membership_level"],
                "parent_first": cancellation_data["parent_first"],
                "parent_last": cancellation_data["parent_last"],
            }

            # Get lesson datetime for policy check
            lesson_datetime = parse_lesson_datetime(
                cancellation_data["lesson_date"], cancellation_data["lesson_time"]
            )

            # Check policy to determine if should be free or charged
            should_be_charged, charge_reason = will_be_charged(
                student_dict, lesson_datetime
            )

            if should_be_charged:
                # Process as charged according to policy
                conn.execute(
                    """UPDATE cancellations 
                       SET charged = 1, status = 'charged', manager_notes = ?, 
                           approved_by = ?, updated_at = ? 
                       WHERE id = ?""",
                    (
                        f"Auto-processed according to policy: {charge_reason}",
                        session.get("user_email", "System"),
                        toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                        cancellation_data["id"],
                    ),
                )
            else:
                # Process as free according to policy
                conn.execute(
                    """UPDATE cancellations 
                       SET status = 'approved', charged = 0, manager_notes = ?, 
                           approved_by = ?, updated_at = ? 
                       WHERE id = ?""",
                    (
                        f"Auto-processed according to policy: {charge_reason}",
                        session.get("user_email", "System"),
                        toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                        cancellation_data["id"],
                    ),
                )

            processed_count += 1

        conn.commit()
        conn.close()

        log_action(
            "process_all_pending",
            f"Processed {processed_count} pending cancellations according to policy",
        )

        return jsonify(
            {
                "success": True,
                "processed": processed_count,
                "message": f"Processed {processed_count} pending cancellations according to policy",
            }
        )

    except Exception as e:
        conn.close()
        return jsonify({"success": False, "message": str(e)})


@app.route("/manager/api/cancellation/revert", methods=["POST"])
@login_required
@admin_required
def revert_cancellation():
    """Revert cancellation back to pending - UPDATED to clear override flags"""
    data = request.json
    cancellation_id = data.get("cancellation_id")

    if not cancellation_id:
        return jsonify({"success": False, "message": "Missing cancellation ID"})

    conn = get_db()

    try:
        conn.execute(
            """UPDATE cancellations 
               SET status = 'pending', charged = 0, excluded = 0, 
                   is_override = 0, manager_notes = ?, updated_at = ? 
               WHERE id = ?""",
            (
                f"Reverted by {session['user_email']} on {toronto_now().strftime('%Y-%m-%d %H:%M')}",
                toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                cancellation_id,
            ),
        )
        conn.commit()
        conn.close()

        log_action(
            "cancellation_reverted",
            f"Cancellation {cancellation_id} reverted to pending",
        )
        return jsonify({"success": True})

    except Exception as e:
        conn.close()
        return jsonify({"success": False, "message": str(e)})


@app.route("/manager/api/cancellation/note", methods=["POST"])
@login_required
@admin_required
def add_cancellation_note():
    """Add note to cancellation - UNCHANGED"""
    data = request.json
    cancellation_id = data.get("cancellation_id")
    note = data.get("note")

    if not cancellation_id or not note:
        return jsonify({"success": False, "message": "Missing required fields"})

    conn = get_db()

    try:
        # Get existing notes
        existing = conn.execute(
            "SELECT manager_notes FROM cancellations WHERE id = ?", (cancellation_id,)
        ).fetchone()

        existing_notes = (
            existing["manager_notes"] if existing and existing["manager_notes"] else ""
        )
        new_notes = f"{existing_notes}\n[{toronto_now().strftime('%Y-%m-%d %H:%M')} - {session['user_email']}]: {note}".strip()

        conn.execute(
            "UPDATE cancellations SET manager_notes = ?, updated_at = ? WHERE id = ?",
            (new_notes, toronto_now().strftime("%Y-%m-%d %H:%M:%S"), cancellation_id),
        )
        conn.commit()
        conn.close()

        log_action("note_added", f"Note added to cancellation {cancellation_id}")
        return jsonify({"success": True})

    except Exception as e:
        conn.close()
        return jsonify({"success": False, "message": str(e)})


@app.route("/manager/api/cancellation/process-all-pending", methods=["POST"])
@login_required
@admin_required
def process_all_pending():
    """Auto-process all pending cancellations according to policy - UPDATED"""
    conn = get_db()

    try:
        # Get all pending cancellations with student info
        pending = conn.execute(
            """
            SELECT c.*, s.membership_level 
            FROM cancellations c
            JOIN students s ON c.student_id = s.id
            WHERE c.status = 'pending' OR c.status IS NULL
        """
        ).fetchall()

        processed_count = 0

        for cancellation in pending:
            # Get membership tier info
            tier = get_membership_tier(cancellation["membership_level"])
            if not tier:
                continue

            # Check if within deadline using existing deadline_passed field or calculate it
            deadline_passed = cancellation.get("deadline_passed", 0)

            if not deadline_passed:
                # Calculate if not already set
                try:
                    lesson_datetime = datetime.strptime(
                        f"{cancellation['lesson_date']} {cancellation['lesson_time']}",
                        "%Y-%m-%d %H:%M:%S",
                    )
                    created_datetime = datetime.strptime(
                        cancellation["created_at"], "%Y-%m-%d %H:%M:%S"
                    )
                    hours_diff = (
                        lesson_datetime - created_datetime
                    ).total_seconds() / 3600
                    deadline_passed = hours_diff < tier["deadline_hours"]
                except:
                    deadline_passed = False

            # Check monthly usage
            monthly_count = get_monthly_cancellation_count(cancellation["student_id"])

            # Determine if should be charged
            should_charge = False
            charge_reason = ""

            if deadline_passed:
                should_charge = True
                charge_reason = "Submitted after deadline"
            elif monthly_count >= tier["free_notices"]:
                should_charge = True
                charge_reason = "Monthly free cancellation limit exceeded"

            # Update cancellation - NOT marked as override since this is automatic
            if should_charge:
                conn.execute(
                    """UPDATE cancellations 
                       SET charged = 1, status = 'charged', manager_notes = ?, 
                           deadline_passed = ?, updated_at = ? 
                       WHERE id = ?""",
                    (
                        f"Auto-processed: {charge_reason}",
                        (
                            1
                            if deadline_passed
                            else cancellation.get("deadline_passed", 0)
                        ),
                        toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                        cancellation["id"],
                    ),
                )
            else:
                conn.execute(
                    """UPDATE cancellations 
                       SET status = 'approved', charged = 0, manager_notes = ?, 
                           deadline_passed = ?, updated_at = ? 
                       WHERE id = ?""",
                    (
                        "Auto-processed: Within policy",
                        (
                            1
                            if deadline_passed
                            else cancellation.get("deadline_passed", 0)
                        ),
                        toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                        cancellation["id"],
                    ),
                )

            processed_count += 1

        conn.commit()
        conn.close()

        log_action(
            "auto_process_all",
            f"Auto-processed {processed_count} pending cancellations",
        )
        return jsonify({"success": True, "processed": processed_count})

    except Exception as e:
        conn.close()
        return jsonify({"success": False, "message": str(e)})


@app.route("/manager/api/student/<int:student_id>/details", methods=["GET"])
@login_required
@admin_required
def get_student_details(student_id):
    """Get detailed student information including cancellation statistics"""
    conn = get_db()

    try:
        # Get student info
        student = conn.execute(
            "SELECT * FROM students WHERE id = ?", (student_id,)
        ).fetchone()

        if not student:
            return jsonify({"success": False, "message": "Student not found"})

        # Get cancellation statistics with more detail
        stats = conn.execute(
            """
            SELECT 
                COUNT(*) as total_cancellations,
                SUM(CASE WHEN strftime('%Y-%m', created_at) = strftime('%Y-%m', 'now') THEN 1 ELSE 0 END) as this_month,
                SUM(CASE WHEN charged = 0 AND excluded = 0 AND strftime('%Y-%m', created_at) = strftime('%Y-%m', 'now') THEN 1 ELSE 0 END) as free_used,
                SUM(CASE WHEN charged = 1 THEN 1 ELSE 0 END) as total_charged,
                SUM(CASE WHEN excluded = 1 THEN 1 ELSE 0 END) as total_excluded,
                SUM(CASE WHEN charged = 1 AND strftime('%Y-%m', created_at) = strftime('%Y-%m', 'now') THEN 1 ELSE 0 END) as charged_this_month,
                SUM(CASE WHEN excluded = 1 AND strftime('%Y-%m', created_at) = strftime('%Y-%m', 'now') THEN 1 ELSE 0 END) as excluded_this_month
            FROM cancellations 
            WHERE student_id = ?
        """,
            (student_id,),
        ).fetchone()

        # Get membership tier info
        tier = get_membership_tier(student["membership_level"])
        monthly_limit = tier["free_notices"] if tier else 1

        # Get recent cancellations
        recent_cancellations = conn.execute(
            """
            SELECT lesson_date, lesson_time, created_at, charged, excluded, status
            FROM cancellations 
            WHERE student_id = ? 
            ORDER BY created_at DESC 
            LIMIT 5
        """,
            (student_id,),
        ).fetchall()

        conn.close()

        return jsonify(
            {
                "success": True,
                "student": dict(student),
                "stats": {
                    "total_cancellations": stats["total_cancellations"] or 0,
                    "this_month": stats["this_month"] or 0,
                    "free_used": stats["free_used"] or 0,
                    "monthly_limit": monthly_limit,
                    "remaining": max(0, monthly_limit - (stats["free_used"] or 0)),
                    "total_charged": stats["total_charged"] or 0,
                    "total_excluded": stats["total_excluded"] or 0,
                    "charged_this_month": stats["charged_this_month"] or 0,
                    "excluded_this_month": stats["excluded_this_month"] or 0,
                },
                "recent_cancellations": [dict(c) for c in recent_cancellations],
            }
        )

    except Exception as e:
        conn.close()
        return jsonify({"success": False, "message": str(e)})


@app.route("/manager/api/cancellation/<int:cancellation_id>/details", methods=["GET"])
@login_required
@admin_required
def get_cancellation_details(cancellation_id):
    """Get detailed cancellation information - UPDATED with new fields"""
    conn = get_db()

    try:
        # Get cancellation with student info - UPDATED query
        cancellation = conn.execute(
            """
            SELECT c.*, s.first_name, s.last_name, s.parent_first, s.parent_last, 
                   s.email, s.phone, s.membership_level, s.created_at as student_created_at
            FROM cancellations c
            JOIN students s ON c.student_id = s.id
            WHERE c.id = ?
        """,
            (cancellation_id,),
        ).fetchone()

        if not cancellation:
            return jsonify({"success": False, "message": "Cancellation not found"})

        # Convert Row to dict to use .get() method
        cancellation_dict = dict(cancellation)

        # Get membership tier info for policy analysis
        tier = get_membership_tier(cancellation_dict["membership_level"])

        # Calculate policy compliance
        try:
            lesson_datetime = datetime.strptime(
                f"{cancellation_dict['lesson_date']} {cancellation_dict['lesson_time']}",
                "%Y-%m-%d %H:%M:%S",
            )
        except ValueError:
            # Handle time without seconds
            lesson_datetime = datetime.strptime(
                f"{cancellation_dict['lesson_date']} {cancellation_dict['lesson_time']}:00",
                "%Y-%m-%d %H:%M:%S",
            )

        created_datetime = datetime.strptime(
            cancellation_dict["created_at"], "%Y-%m-%d %H:%M:%S"
        )
        hours_notice = (lesson_datetime - created_datetime).total_seconds() / 3600

        # Get monthly usage
        monthly_count = get_monthly_cancellation_count(cancellation_dict["student_id"])

        # Parse sequential lessons
        sequential_lessons = []
        if cancellation_dict["sequential_lessons"]:
            try:
                sequential_lessons = eval(cancellation_dict["sequential_lessons"])
            except:
                sequential_lessons = []

        # Get action history from logs
        action_history = conn.execute(
            """
            SELECT action, details, created_at
            FROM system_logs
            WHERE details LIKE ?
            ORDER BY created_at DESC
            LIMIT 10
        """,
            (f"%{cancellation_id}%",),
        ).fetchall()

        conn.close()

        # Use deadline_passed from database if available, otherwise calculate
        within_deadline = True
        if cancellation_dict.get("deadline_passed") is not None:
            within_deadline = not bool(cancellation_dict["deadline_passed"])
        elif tier:
            within_deadline = hours_notice >= tier["deadline_hours"]

        return jsonify(
            {
                "success": True,
                "cancellation": {
                    **cancellation_dict,
                    "sequential_lessons": sequential_lessons,
                },
                "student": {
                    "first_name": cancellation_dict["first_name"],
                    "last_name": cancellation_dict["last_name"],
                    "parent_first": cancellation_dict["parent_first"] or "",
                    "parent_last": cancellation_dict["parent_last"] or "",
                    "email": cancellation_dict["email"],
                    "phone": cancellation_dict["phone"] or "",
                    "membership_level": cancellation_dict["membership_level"],
                },
                "policy": {
                    "monthly_limit": tier["free_notices"] if tier else 1,
                    "used_this_month": monthly_count,
                    "remaining": max(
                        0, (tier["free_notices"] if tier else 1) - monthly_count
                    ),
                    "deadline_display": (
                        tier["deadline_display"] if tier else "6pm previous day"
                    ),
                    "hours_notice": round(hours_notice, 1),
                    "within_deadline": within_deadline,
                    "policy_result": (
                        "Within policy"
                        if within_deadline
                        and monthly_count < (tier["free_notices"] if tier else 1)
                        else "Policy violation"
                    ),
                },
                "action_history": [dict(a) for a in action_history],
            }
        )

    except Exception as e:
        conn.close()
        print(f"Error in get_cancellation_details: {str(e)}")
        return jsonify({"success": False, "message": str(e)})


# Update the existing manager_cancellations route to handle student and view parameters properly
# Debug version of the manager_cancellations route
@app.route("/manager/cancellations")
@login_required
@admin_required
def manager_cancellations():
    """Manager cancellations page - UPDATED with consistent date processing"""
    # Get filter parameters
    filter_status = request.args.get("status", "")
    search = request.args.get("search", "")
    date_range = request.args.get("date_range", "month")
    membership = request.args.get("membership", "")
    sort_by = request.args.get("sort", "submit_date")
    student_id = request.args.get("student")
    date_from = request.args.get("date_from")
    date_to = request.args.get("date_to")

    # Check if this is an AJAX request
    is_ajax = request.headers.get("X-Requested-With") == "XMLHttpRequest"

    conn = get_db()

    # Build query based on filters
    where_clauses = []
    params = []

    # Add student filter
    if student_id:
        where_clauses.append("c.student_id = ?")
        params.append(student_id)

    # Status filters - UPDATED with new options
    if filter_status == "free":
        where_clauses.append("c.charged = 0 AND c.status = 'approved'")
    elif filter_status == "charged":
        where_clauses.append("c.charged = 1")
    elif filter_status == "excluded":
        where_clauses.append("c.excluded = 1")
    elif filter_status == "note":
        where_clauses.append(
            "c.cancellation_note IS NOT NULL AND c.cancellation_note != ''"
        )
    elif filter_status == "deadline_passed":
        where_clauses.append("c.deadline_passed = 1")
    elif filter_status == "override":
        where_clauses.append("c.is_override = 1")

    # Search filter
    if search:
        where_clauses.append(
            "(s.first_name LIKE ? OR s.last_name LIKE ? OR s.email LIKE ?)"
        )
        search_param = f"%{search}%"
        params.extend([search_param, search_param, search_param])

    # Date range filter - UPDATED with new options
    if date_range == "today":
        where_clauses.append("DATE(c.created_at) = DATE('now')")
    elif date_range == "yesterday":
        where_clauses.append("DATE(c.created_at) = DATE('now', '-1 day')")
    elif date_range == "7days":
        where_clauses.append("c.created_at >= DATE('now', '-7 days')")
    elif date_range == "month":
        where_clauses.append("c.created_at >= DATE('now', '-30 days')")
    elif date_range == "all":
        pass  # No date filter
    elif date_range == "custom" and date_from and date_to:
        where_clauses.append("DATE(c.created_at) BETWEEN ? AND ?")
        params.extend([date_from, date_to])

    # Membership filter
    if membership:
        where_clauses.append("s.membership_level = ?")
        params.append(membership)

    where_sql = " AND ".join(where_clauses) if where_clauses else "1=1"

    # Sort order - UPDATED with new options
    if sort_by == "submit_date":
        order_by = "c.created_at DESC"
    elif sort_by == "lesson_date":
        order_by = "c.lesson_date DESC"
    elif sort_by == "student":
        order_by = "s.last_name, s.first_name"
    elif sort_by == "status":
        order_by = "c.status, c.charged, c.excluded"
    else:
        order_by = "c.created_at DESC"

    # Get cancellations with student info
    cancellations_raw = conn.execute(
        f"""
        SELECT c.*, s.first_name, s.last_name, s.email, s.membership_level
        FROM cancellations c
        JOIN students s ON c.student_id = s.id
        WHERE {where_sql}
        ORDER BY {order_by}
        LIMIT 100
        """,
        params,
    ).fetchall()

    # Process cancellations data with consistent date formatting
    cancellations = []
    for row in cancellations_raw:
        cancellation = process_cancellation_dates(row)

        # Add computed fields
        cancellation["student_name"] = (
            f"{cancellation['first_name']} {cancellation['last_name']}"
        )

        # Status class for CSS - UPDATED with new status types
        if cancellation.get("cancellation_note"):
            cancellation["status_class"] = "note"
        elif cancellation.get("deadline_passed"):
            cancellation["status_class"] = "deadline-passed"
        elif cancellation.get("is_override"):
            cancellation["status_class"] = "override"
        elif cancellation["excluded"]:
            cancellation["status_class"] = "excluded"
        elif cancellation["charged"]:
            cancellation["status_class"] = "charged"
        elif cancellation.get("status") == "approved":
            cancellation["status_class"] = "free"
        else:
            cancellation["status_class"] = "pending"

        # Calculate deadline status
        tier = get_membership_tier(cancellation["membership_level"])
        if tier and cancellation.get("created_at"):
            lesson_datetime = datetime.combine(
                cancellation["lesson_date"], cancellation["lesson_time"]
            )
            hours_notice = (
                lesson_datetime - cancellation["created_at"]
            ).total_seconds() / 3600
            cancellation["within_deadline"] = hours_notice >= tier["deadline_hours"]
        else:
            cancellation["within_deadline"] = True

        # Add usage statistics
        cancellation["used_this_month"] = get_monthly_cancellation_count(
            cancellation.get("student_id")
        )
        cancellation["monthly_limit"] = tier["free_notices"] if tier else 1

        # Additional fields
        cancellation["reschedule_requested"] = bool(
            cancellation.get("reschedule_requested")
        )
        cancellation["reschedule_preferences"] = cancellation.get(
            "reschedule_preferences", ""
        )
        cancellation["error_report"] = cancellation.get("error_report", "")
        cancellation["approved_by"] = cancellation.get("approved_by", "")

        # Add urgency flags
        if cancellation.get("created_at"):
            hours_since = (
                datetime.now() - cancellation["created_at"]
            ).total_seconds() / 3600
            cancellation["is_recent"] = hours_since < 2
            cancellation["is_urgent"] = (
                hours_since > 24 and cancellation.get("status") == "pending"
            )
        else:
            cancellation["is_recent"] = False
            cancellation["is_urgent"] = False

        cancellations.append(cancellation)

    # Get summary stats
    stats_raw = conn.execute(
        f"""
        SELECT
            COUNT(*) as total_cancellations,
            SUM(CASE WHEN DATE(c.created_at) = DATE('now') THEN 1 ELSE 0 END) as today_cancellations,
            SUM(CASE WHEN c.charged = 0 AND c.status = 'approved' THEN 1 ELSE 0 END) as free_cancellations,
            SUM(CASE WHEN c.charged = 1 THEN 1 ELSE 0 END) as charged_cancellations,
            SUM(CASE WHEN c.excluded = 1 THEN 1 ELSE 0 END) as excluded_cancellations,
            SUM(CASE WHEN c.is_override = 1 AND strftime('%Y-%m', c.created_at) = strftime('%Y-%m', 'now') THEN 1 ELSE 0 END) as override_this_month,
            SUM(CASE WHEN c.deadline_passed = 1 AND c.created_at >= DATE('now', '-7 days') THEN 1 ELSE 0 END) as deadline_passed,
            SUM(CASE WHEN c.cancellation_note IS NOT NULL AND c.cancellation_note != '' AND c.created_at >= DATE('now', '-7 days') THEN 1 ELSE 0 END) as with_notes
        FROM cancellations c
        WHERE {where_sql}
        """,
        params,
    ).fetchone()

    stats = (
        dict(stats_raw)
        if stats_raw
        else {
            "total_cancellations": 0,
            "today_cancellations": 0,
            "free_cancellations": 0,
            "charged_cancellations": 0,
            "excluded_cancellations": 0,
            "override_this_month": 0,
            "deadline_passed": 0,
            "with_notes": 0,
        }
    )

    # Get membership tiers for filter dropdown
    membership_tiers = [
        "Bronze",
        "Silver",
        "Gold",
        "Intro Package",
        "Legacy",
        "Guest",
        "Welcome Package",
    ]

    # Handle student context for filtering
    current_student_id = None
    filtered_student_name = ""

    if student_id:
        try:
            student_info = conn.execute(
                "SELECT first_name, last_name FROM students WHERE id = ?", (student_id,)
            ).fetchone()
            if student_info:
                current_student_id = int(student_id)
                filtered_student_name = (
                    f"{student_info['first_name']} {student_info['last_name']}"
                )
        except (ValueError, TypeError):
            pass

    # Mock pagination
    pagination = type(
        "Pagination",
        (),
        {
            "pages": 1,
            "has_prev": False,
            "has_next": False,
            "prev_num": None,
            "next_num": None,
            "page": 1,
            "iter_pages": lambda: [1],
        },
    )()

    conn.close()

    return render_template(
        "manager_cancellations.html",
        cancellations=cancellations,
        stats=stats,
        filter_status=filter_status,
        membership_tiers=membership_tiers,
        pagination=pagination,
        current_student_id=current_student_id,
        filtered_student_name=filtered_student_name,
    )


# Add these helper functions to your app.py file in the UTILITY FUNCTIONS section


def prepare_cancellations_for_json(cancellations):
    """
    Prepare cancellation data specifically for JSON serialization
    """
    json_cancellations = []
    for cancellation in cancellations:
        # Create a clean dict with only serializable data
        json_cancellation = {
            "id": cancellation.get("id"),
            "student_name": cancellation.get("student_name", ""),
            "first_name": cancellation.get("first_name", ""),
            "last_name": cancellation.get("last_name", ""),
            "membership_level": cancellation.get("membership_level", ""),
            "lesson_date": str(cancellation.get("lesson_date", "")),
            "lesson_time": str(cancellation.get("lesson_time", "")),
            "created_at": str(cancellation.get("created_at", "")),
            "status": cancellation.get("status", ""),
            "charged": bool(cancellation.get("charged", False)),
            "excluded": bool(cancellation.get("excluded", False)),
            "deadline_passed": bool(cancellation.get("deadline_passed", False)),
            "is_override": bool(cancellation.get("is_override", False)),
            "cancellation_note": str(cancellation.get("cancellation_note", "") or ""),
            "manager_notes": str(cancellation.get("manager_notes", "") or ""),
            "reschedule_requested": bool(
                cancellation.get("reschedule_requested", False)
            ),
            "reschedule_preferences": str(
                cancellation.get("reschedule_preferences", "") or ""
            ),
            "error_report": str(cancellation.get("error_report", "") or ""),
        }
        json_cancellations.append(json_cancellation)
    return json_cancellations


def make_json_serializable(obj):
    """
    Convert objects to JSON serializable format
    """
    from datetime import datetime, date, time

    if isinstance(obj, dict):
        return {key: make_json_serializable(value) for key, value in obj.items()}
    elif isinstance(obj, list):
        return [make_json_serializable(item) for item in obj]
    elif isinstance(obj, (date, time, datetime)):
        return str(obj)
    elif hasattr(obj, "__dict__"):
        return make_json_serializable(obj.__dict__)
    else:
        return obj


def get_analytics_summary(
    date_range="month",
    status_filter="",
    membership_filter="",
    start_date="",
    end_date="",
):
    """
    Get analytics summary for dashboard widgets
    """
    conn = get_db()

    try:
        # Build filters similar to main analytics
        date_conditions = {
            "today": "DATE(c.created_at) = DATE('now')",
            "7days": "c.created_at >= DATE('now', '-7 days')",
            "month": "c.created_at >= DATE('now', '-30 days')",
            "all": "1=1",
        }

        # Handle custom date range
        if date_range == "custom" and start_date and end_date:
            date_clause = "DATE(c.created_at) BETWEEN ? AND ?"
        else:
            date_clause = date_conditions.get(date_range, date_conditions["month"])

        where_clause = f"WHERE {date_clause}"
        params = []

        # Add custom date parameters if needed
        if date_range == "custom" and start_date and end_date:
            params.extend([start_date, end_date])

        if status_filter:
            status_conditions = {
                "free": "c.charged = 0 AND c.excluded = 0 AND c.status = 'approved'",
                "charged": "c.charged = 1 AND c.excluded = 0",
                "excluded": "c.excluded = 1",
                "deadline_passed": "c.deadline_passed = 1",
                "override": "c.is_override = 1",
            }
            if status_filter in status_conditions:
                where_clause += f" AND {status_conditions[status_filter]}"

        if membership_filter:
            where_clause += " AND s.membership_level = ?"
            params.append(membership_filter)

        summary_query = f"""
            SELECT
                COUNT(*) as total_count,
                COUNT(DISTINCT c.student_id) as unique_students,
                AVG(CASE WHEN c.charged = 1 THEN 1.0 ELSE 0.0 END) as charge_rate,
                COUNT(CASE WHEN c.is_override = 1 THEN 1 END) as override_count
            FROM cancellations c
            JOIN students s ON c.student_id = s.id
            {where_clause}
        """

        result = conn.execute(summary_query, params).fetchone()
        conn.close()

        return {
            "total_count": result["total_count"] or 0,
            "unique_students": result["unique_students"] or 0,
            "charge_rate": round((result["charge_rate"] or 0) * 100, 1),
            "override_count": result["override_count"] or 0,
        }

    except Exception as e:
        conn.close()
        print(f"Error in get_analytics_summary: {e}")
        return {
            "total_count": 0,
            "unique_students": 0,
            "charge_rate": 0.0,
            "override_count": 0,
        }


# Replace your manager_analytics route with this version that fixes the data structure issues


@app.route("/manager/analytics")
@login_required
@admin_required
def manager_analytics():
    """Manager analytics page with consistent data between stats and charts"""
    # Get filter parameters
    date_range = request.args.get("date_range", "month")
    status_filter = request.args.get("status", "")
    membership_filter = request.args.get("membership", "")
    start_date = request.args.get("start_date", "")
    end_date = request.args.get("end_date", "")
    format_type = request.args.get("format", "html")

    conn = get_db()

    try:
        print(f"\n=== ANALYTICS DEBUG START ===")
        print(
            f"Filters: date_range={date_range}, status={status_filter}, membership={membership_filter}"
        )

        # FIRST: Let's see what data we actually have
        total_in_db = conn.execute(
            "SELECT COUNT(*) as count FROM cancellations"
        ).fetchone()["count"]
        print(f"Total cancellations in database: {total_in_db}")

        # Check date range of actual data
        date_range_info = conn.execute(
            """
            SELECT 
                MIN(created_at) as earliest,
                MAX(created_at) as latest,
                COUNT(*) as total
            FROM cancellations
        """
        ).fetchone()

        if date_range_info:
            print(
                f"Data spans: {date_range_info['earliest']} to {date_range_info['latest']}"
            )

        # Build date filter for FILTERED data (stats)
        date_conditions = {
            "today": "DATE(c.created_at) = DATE('now')",
            "7days": "c.created_at >= DATE('now', '-7 days')",
            "month": "c.created_at >= DATE('now', '-30 days')",
            "all": "1=1",  # No date filter - this is what finds all 14!
        }

        # Handle custom date range
        if date_range == "custom" and start_date and end_date:
            date_clause = "DATE(c.created_at) BETWEEN ? AND ?"
        else:
            date_clause = date_conditions.get(date_range, date_conditions["month"])

        # Build status filter
        status_conditions = {
            "free": "c.charged = 0 AND c.excluded = 0 AND c.status = 'approved'",
            "charged": "c.charged = 1 AND c.excluded = 0",
            "excluded": "c.excluded = 1",
            "deadline_passed": "c.deadline_passed = 1",
            "override": "c.is_override = 1",
            "note": "c.cancellation_note IS NOT NULL AND c.cancellation_note != ''",
        }
        status_clause = status_conditions.get(status_filter, "1=1")

        # Build membership filter
        where_clause = f"WHERE {date_clause} AND {status_clause}"
        params = []

        # Add custom date parameters if needed
        if date_range == "custom" and start_date and end_date:
            params.extend([start_date, end_date])

        if membership_filter:
            where_clause += " AND s.membership_level = ?"
            params.append(membership_filter)

        print(f"Filter WHERE clause: {where_clause}")
        print(f"Filter params: {params}")

        # 1. Summary statistics with applied filters
        stats_query = f"""
            SELECT
                CAST(COUNT(*) AS INTEGER) as total_cancellations,
                CAST(SUM(CASE WHEN c.charged = 0 AND c.excluded = 0 AND c.status = 'approved' THEN 1 ELSE 0 END) AS INTEGER) as free_cancellations,
                CAST(SUM(CASE WHEN c.charged = 1 AND c.excluded = 0 THEN 1 ELSE 0 END) AS INTEGER) as charged_cancellations,
                CAST(SUM(CASE WHEN c.excluded = 1 THEN 1 ELSE 0 END) AS INTEGER) as excluded_cancellations,
                CAST(SUM(CASE WHEN c.is_override = 1 THEN 1 ELSE 0 END) AS INTEGER) as override_cancellations,
                CAST(SUM(CASE WHEN c.deadline_passed = 1 THEN 1 ELSE 0 END) AS INTEGER) as deadline_passed,
                CAST(SUM(CASE WHEN c.cancellation_note IS NOT NULL AND c.cancellation_note != '' THEN 1 ELSE 0 END) AS INTEGER) as with_notes
            FROM cancellations c
            JOIN students s ON c.student_id = s.id
            {where_clause}
        """

        stats_raw = conn.execute(stats_query, params).fetchone()

        # Process stats
        stats = {}
        if stats_raw:
            for key in stats_raw.keys():
                value = stats_raw[key]
                stats[key] = int(value) if value is not None else 0
        else:
            stats = {
                "total_cancellations": 0,
                "free_cancellations": 0,
                "charged_cancellations": 0,
                "excluded_cancellations": 0,
                "override_cancellations": 0,
                "deadline_passed": 0,
                "with_notes": 0,
            }

        print(f"STATS with current filters: {stats}")

        # 2. Monthly trends - USE SAME FILTERS for consistency!
        monthly_trends_query = f"""
            SELECT 
                strftime('%Y-%m', c.created_at) as month,
                CAST(COUNT(*) AS INTEGER) as total_cancellations,
                CAST(SUM(CASE WHEN c.charged = 0 AND c.excluded = 0 AND c.status = 'approved' THEN 1 ELSE 0 END) AS INTEGER) as free_cancellations,
                CAST(SUM(CASE WHEN c.charged = 1 AND c.excluded = 0 THEN 1 ELSE 0 END) AS INTEGER) as charged_cancellations,
                CAST(SUM(CASE WHEN c.is_override = 1 THEN 1 ELSE 0 END) AS INTEGER) as override_cancellations
            FROM cancellations c
            JOIN students s ON c.student_id = s.id
            {where_clause}
            GROUP BY strftime('%Y-%m', c.created_at)
            ORDER BY month ASC
        """

        print("Monthly trends query with same filters:")
        print(monthly_trends_query)
        print(f"With params: {params}")

        monthly_trends_raw = conn.execute(monthly_trends_query, params).fetchall()
        monthly_trends = []

        print(f"Monthly trends raw results: {len(monthly_trends_raw)} months")

        if monthly_trends_raw:
            total_from_trends = 0
            for trend in monthly_trends_raw:
                try:
                    month_date = datetime.strptime(trend["month"], "%Y-%m")
                    trend_data = {
                        "month": month_date.strftime("%b %Y"),
                        "month_short": month_date.strftime("%m/%y"),
                        "total_cancellations": int(trend["total_cancellations"] or 0),
                        "free_cancellations": int(trend["free_cancellations"] or 0),
                        "charged_cancellations": int(
                            trend["charged_cancellations"] or 0
                        ),
                        "override_cancellations": int(
                            trend["override_cancellations"] or 0
                        ),
                    }
                    monthly_trends.append(trend_data)
                    total_from_trends += trend_data["total_cancellations"]
                    print(
                        f"  {trend_data['month']}: {trend_data['total_cancellations']} total"
                    )
                except (ValueError, TypeError) as e:
                    print(f"Error parsing month {trend['month']}: {e}")
                    continue

            print(f"Total from trends: {total_from_trends}")
            print(f"Total from stats: {stats['total_cancellations']}")

            if total_from_trends != stats["total_cancellations"]:
                print("⚠️  WARNING: Trends total doesn't match stats total!")

        # If no monthly data found, create structure based on actual data
        if not monthly_trends:
            print("No monthly trends found, checking if any cancellations exist...")

            if stats["total_cancellations"] > 0:
                print("Stats show cancellations exist, but trends query found none!")

                # Let's see what months actually have data (no filters)
                all_months_query = """
                    SELECT DISTINCT 
                        strftime('%Y-%m', created_at) as month,
                        COUNT(*) as count
                    FROM cancellations 
                    GROUP BY strftime('%Y-%m', created_at)
                    ORDER BY month DESC
                """
                all_months = conn.execute(all_months_query).fetchall()
                print(
                    f"All months in DB: {[(m['month'], m['count']) for m in all_months]}"
                )

                # Create trend for current month with aggregated data
                current_month = datetime.now()
                monthly_trends = [
                    {
                        "month": current_month.strftime("%b %Y"),
                        "month_short": current_month.strftime("%m/%y"),
                        "total_cancellations": stats["total_cancellations"],
                        "free_cancellations": stats["free_cancellations"],
                        "charged_cancellations": stats["charged_cancellations"],
                        "override_cancellations": stats["override_cancellations"],
                    }
                ]
                print(f"Created synthetic monthly trend: {monthly_trends[0]}")
            else:
                # Truly no data
                current_month = datetime.now()
                monthly_trends = [
                    {
                        "month": current_month.strftime("%b %Y"),
                        "month_short": current_month.strftime("%m/%y"),
                        "total_cancellations": 0,
                        "free_cancellations": 0,
                        "charged_cancellations": 0,
                        "override_cancellations": 0,
                    }
                ]

        # 3. Top students (with same filters)
        top_students_query = f"""
            SELECT 
                s.first_name, s.last_name, s.membership_level, 
                CAST(COUNT(c.id) AS INTEGER) as cancellation_count
            FROM students s
            JOIN cancellations c ON s.id = c.student_id
            {where_clause}
            GROUP BY s.id, s.first_name, s.last_name, s.membership_level
            HAVING cancellation_count > 0
            ORDER BY cancellation_count DESC
            LIMIT 10
        """

        top_students_raw = conn.execute(top_students_query, params).fetchall()
        top_students = [dict(s) for s in top_students_raw]

        # 4. Membership tier analysis (with same filters)
        tier_analysis_query = f"""
            SELECT 
                s.membership_level,
                CAST(COUNT(DISTINCT s.id) AS INTEGER) as student_count,
                CAST(COUNT(c.id) AS INTEGER) as total_cancellations,
                CAST(SUM(CASE WHEN c.charged = 0 AND c.excluded = 0 AND c.status = 'approved' THEN 1 ELSE 0 END) AS INTEGER) as free_cancellations,
                CAST(SUM(CASE WHEN c.charged = 1 AND c.excluded = 0 THEN 1 ELSE 0 END) AS INTEGER) as charged_cancellations,
                CAST(ROUND(COUNT(c.id) * 1.0 / NULLIF(COUNT(DISTINCT s.id), 0), 2) AS REAL) as avg_monthly_cancellations
            FROM students s
            JOIN cancellations c ON s.id = c.student_id
            {where_clause}
            GROUP BY s.membership_level
            HAVING total_cancellations > 0
            ORDER BY student_count DESC
        """

        tier_distribution_raw = conn.execute(tier_analysis_query, params).fetchall()
        tier_distribution = [dict(t) for t in tier_distribution_raw]

        # 5. Filtered cancellations for table (with same filters)
        filtered_cancellations_query = f"""
            SELECT c.*, s.first_name, s.last_name, s.membership_level,
                   (s.first_name || ' ' || s.last_name) as student_name
            FROM cancellations c
            JOIN students s ON c.student_id = s.id
            {where_clause}
            ORDER BY c.created_at DESC
            LIMIT 50
        """

        filtered_cancellations_raw = conn.execute(
            filtered_cancellations_query, params
        ).fetchall()

        # Process cancellations
        filtered_cancellations = []
        for cancellation in filtered_cancellations_raw:
            processed_cancellation = process_cancellation_dates(cancellation)
            filtered_cancellations.append(processed_cancellation)

        # 6. Revenue impact (no filters - always show all for broader context)
        revenue_impact_query = """
            SELECT
                strftime('%Y-%m', c.created_at) as month,
                CAST(SUM(CASE WHEN c.charged = 1 THEN 1 ELSE 0 END) AS INTEGER) as charged_count,
                CAST(SUM(CASE WHEN c.charged = 1 THEN 1 ELSE 0 END) * 25 AS INTEGER) as estimated_revenue
            FROM cancellations c
            WHERE c.created_at >= DATE('now', '-12 months')
            GROUP BY strftime('%Y-%m', c.created_at)
            ORDER BY month ASC
        """

        revenue_impact_raw = conn.execute(revenue_impact_query).fetchall()
        revenue_impact = []

        for revenue in revenue_impact_raw:
            try:
                month_date = datetime.strptime(revenue["month"], "%Y-%m")
                revenue_impact.append(
                    {
                        "month": month_date.strftime("%b %Y"),
                        "charged_count": int(revenue["charged_count"] or 0),
                        "estimated_revenue": int(revenue["estimated_revenue"] or 0),
                    }
                )
            except (ValueError, TypeError):
                continue

        conn.close()

        print(f"FINAL RESULTS:")
        print(f"  Stats: {stats}")
        print(f"  Monthly trends: {len(monthly_trends)} months")
        print(f"  First month: {monthly_trends[0] if monthly_trends else 'None'}")
        print(f"=== ANALYTICS DEBUG END ===\n")

        # Handle JSON request for AJAX
        if format_type == "json":
            json_data = {
                "success": True,
                "stats": stats,
                "monthly_trends": monthly_trends,
                "top_students": top_students,
                "tier_distribution": tier_distribution,
                "filtered_cancellations": prepare_cancellations_for_json(
                    filtered_cancellations
                ),
                "revenue_impact": revenue_impact,
            }
            return jsonify(json_data)

        # Regular HTML response
        log_action(
            "analytics_viewed",
            f"Filters: {date_range}, {status_filter}, {membership_filter}",
        )

        return render_template(
            "manager_analytics.html",
            stats=stats,
            monthly_trends=monthly_trends,
            tier_distribution=tier_distribution,
            top_students=top_students,
            filtered_cancellations=filtered_cancellations,
            revenue_impact=revenue_impact,
            current_filters={
                "date_range": date_range,
                "status": status_filter,
                "membership": membership_filter,
            },
            debug_info={
                "trends_count": len(monthly_trends),
                "has_data": len(monthly_trends) > 0
                and any(t["total_cancellations"] > 0 for t in monthly_trends),
                "first_month": monthly_trends[0]["month"] if monthly_trends else "None",
                "total_in_first": (
                    monthly_trends[0]["total_cancellations"] if monthly_trends else 0
                ),
            },
        )

    except Exception as e:
        conn.close()
        print(f"ERROR in manager_analytics: {str(e)}")
        import traceback

        traceback.print_exc()

        log_action("analytics_error", f"Error loading analytics: {str(e)}")

        # Return safe fallback
        fallback_stats = {
            "total_cancellations": 0,
            "free_cancellations": 0,
            "charged_cancellations": 0,
            "excluded_cancellations": 0,
            "override_cancellations": 0,
            "deadline_passed": 0,
            "with_notes": 0,
        }

        fallback_trends = [
            {
                "month": datetime.now().strftime("%b %Y"),
                "month_short": datetime.now().strftime("%m/%y"),
                "total_cancellations": 0,
                "free_cancellations": 0,
                "charged_cancellations": 0,
                "override_cancellations": 0,
            }
        ]

        if format_type == "json":
            return jsonify(
                {
                    "success": False,
                    "error": str(e),
                    "stats": fallback_stats,
                    "monthly_trends": fallback_trends,
                    "tier_distribution": [],
                    "top_students": [],
                    "filtered_cancellations": [],
                    "revenue_impact": [],
                }
            )

        return render_template(
            "manager_analytics.html",
            stats=fallback_stats,
            monthly_trends=fallback_trends,
            tier_distribution=[],
            top_students=[],
            filtered_cancellations=[],
            revenue_impact=[],
            error_message=f"Analytics error: {str(e)}",
            current_filters={"date_range": "month", "status": "", "membership": ""},
            debug_info={
                "trends_count": 1,
                "has_data": False,
                "first_month": "Error",
                "total_in_first": 0,
            },
        )


# Manager API endpoints
@app.route("/manager/api/student/<int:student_id>", methods=["PUT", "POST"])
@login_required
@admin_required
def update_student(student_id):
    """Update student information"""
    data = request.json

    conn = get_db()
    conn.execute(
        """
        UPDATE students 
        SET first_name = ?, last_name = ?, parent_first = ?, parent_last = ?,
            email = ?, phone = ?, membership_level = ?
        WHERE id = ?
    """,
        (
            data["first_name"],
            data["last_name"],
            data["parent_first"],
            data["parent_last"],
            data["email"],
            data["phone"],
            data["membership_level"],
            student_id,
        ),
    )
    conn.commit()
    conn.close()

    log_action("student_updated", f"Student ID: {student_id}")
    return jsonify({"success": True})


@app.route("/manager/api/student", methods=["POST"])
@login_required
@admin_required
def add_student():
    """Add new student"""
    data = request.json

    conn = get_db()
    try:
        cursor = conn.execute(
            """
            INSERT INTO students (first_name, last_name, parent_first, parent_last, email, phone, membership_level)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
            (
                data["first_name"],
                data["last_name"],
                data["parent_first"],
                data["parent_last"],
                data["email"],
                data["phone"],
                data["membership_level"],
            ),
        )
        student_id = cursor.lastrowid
        conn.commit()
        conn.close()

        log_action("student_added", f"New student: {data['email']}")
        return jsonify({"success": True, "student_id": student_id})
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({"success": False, "error": "Email already exists"})


@app.route("/manager/api/student/<int:student_id>", methods=["DELETE"])
@login_required
@admin_required
def delete_student(student_id):
    """Delete student"""
    conn = get_db()

    # Get student info for logging
    student = conn.execute(
        "SELECT email FROM students WHERE id = ?", (student_id,)
    ).fetchone()

    # Delete cancellations first (foreign key constraint)
    conn.execute("DELETE FROM cancellations WHERE student_id = ?", (student_id,))
    conn.execute("DELETE FROM students WHERE id = ?", (student_id,))
    conn.commit()
    conn.close()

    log_action(
        "student_deleted",
        f"Student ID: {student_id}, Email: {student['email'] if student else 'Unknown'}",
    )
    return jsonify({"success": True})


@app.route("/manager/api/cancellation/<int:cancellation_id>/exclude", methods=["POST"])
@login_required
@admin_required
def exclude_cancellation(cancellation_id):
    """Exclude cancellation from policy (illness, etc.) - UPDATED with override tracking"""
    data = request.json
    reason = data.get("reason", "")

    if not reason.strip():
        return jsonify({"success": False, "message": "Exclusion reason is required"})

    conn = get_db()

    try:
        conn.execute(
            """UPDATE cancellations 
               SET excluded = 1, exclusion_reason = ?, approved_by = ?, 
                   is_override = 1, updated_at = ? 
               WHERE id = ?""",
            (
                reason,
                session["user_email"],
                toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                cancellation_id,
            ),
        )
        conn.commit()
        conn.close()

        log_action(
            "cancellation_excluded",
            f"Cancellation ID: {cancellation_id}, Reason: {reason}, By: {session['user_email']}",
        )
        return jsonify({"success": True})

    except Exception as e:
        conn.close()
        return jsonify({"success": False, "message": str(e)})


@app.route("/manager/api/bulk-import", methods=["POST"])
@login_required
@admin_required
def bulk_import_students():
    """Bulk import students from CSV data"""
    data = request.json
    csv_data = data.get("csv_data", "")

    if not csv_data:
        return jsonify({"success": False, "error": "No data provided"})

    # Parse CSV data
    csv_reader = csv.reader(io.StringIO(csv_data))
    imported = 0
    errors = []

    conn = get_db()

    for row_num, row in enumerate(csv_reader, 1):
        if len(row) < 6:  # Minimum required columns
            errors.append(f"Row {row_num}: Insufficient data")
            continue

        try:
            conn.execute(
                """
                INSERT INTO students (first_name, last_name, parent_first, parent_last, email, phone, membership_level)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    row[0],
                    row[1],
                    row[2],
                    row[3],
                    row[4],
                    row[5],
                    row[6] if len(row) > 6 else "Bronze",
                ),
            )
            imported += 1
        except sqlite3.IntegrityError:
            errors.append(f"Row {row_num}: Email {row[4]} already exists")
        except Exception as e:
            errors.append(f"Row {row_num}: {str(e)}")

    conn.commit()
    conn.close()

    log_action("bulk_import", f"Imported {imported} students, {len(errors)} errors")
    return jsonify({"success": True, "imported": imported, "errors": errors})


@app.route("/manager/export/students")
@login_required
@admin_required
def export_students():
    """Export students to CSV"""
    conn = get_db()
    students = conn.execute(
        "SELECT * FROM students ORDER BY last_name, first_name"
    ).fetchall()
    conn.close()

    # Create CSV output
    output = io.StringIO()
    writer = csv.writer(output)

    # Header
    writer.writerow(
        [
            "First Name",
            "Last Name",
            "Parent First",
            "Parent Last",
            "Email",
            "Phone",
            "Membership Level",
            "Created",
        ]
    )

    # Data rows
    for student in students:
        writer.writerow(
            [
                student["first_name"],
                student["last_name"],
                student["parent_first"],
                student["parent_last"],
                student["email"],
                student["phone"],
                student["membership_level"],
                student["created_at"],
            ]
        )

    # Create response
    response = app.response_class(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=students_export.csv"},
    )

    log_action("export_students", f"Exported {len(students)} students")
    return response


@app.route("/manager/api/student/<int:student_id>", methods=["POST"])
@login_required
@admin_required
def update_student_post(student_id):
    """Update student information via POST (for the edit form)"""
    try:
        data = request.json
        conn = get_db()

        conn.execute(
            """
            UPDATE students 
            SET first_name = ?, last_name = ?, parent_first = ?, parent_last = ?,
                email = ?, phone = ?, membership_level = ?, updated_at = ?
            WHERE id = ?
        """,
            (
                data["first_name"],
                data["last_name"],
                data.get("parent_first", ""),
                data.get("parent_last", ""),
                data["email"],
                data.get("phone", ""),
                data["membership_level"],
                toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                student_id,
            ),
        )
        conn.commit()
        conn.close()

        log_action("student_updated", f"Student ID: {student_id}")
        return jsonify({"success": True, "message": "Student updated successfully"})

    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


@app.route("/manager/api/student/<int:student_id>/delete", methods=["DELETE"])
@login_required
@admin_required
def delete_student_api(student_id):
    """Delete student via API"""
    try:
        conn = get_db()

        # Get student info for logging
        student = conn.execute(
            "SELECT email, first_name, last_name FROM students WHERE id = ?",
            (student_id,),
        ).fetchone()

        if not student:
            return jsonify({"success": False, "message": "Student not found"})

        # Delete cancellations first (foreign key constraint)
        conn.execute("DELETE FROM cancellations WHERE student_id = ?", (student_id,))
        conn.execute("DELETE FROM students WHERE id = ?", (student_id,))
        conn.commit()
        conn.close()

        log_action(
            "student_deleted",
            f"Student: {student['first_name']} {student['last_name']} ({student['email']})",
        )
        return jsonify({"success": True, "message": "Student deleted successfully"})

    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


@app.route("/manager/api/student/<int:student_id>/membership", methods=["POST"])
@login_required
@admin_required
def change_student_membership(student_id):
    """Change student membership level"""
    try:
        data = request.json
        new_membership = data.get("membership_level")

        if not new_membership:
            return jsonify(
                {"success": False, "message": "Membership level is required"}
            )

        conn = get_db()

        # Get student info for logging
        student = conn.execute(
            "SELECT first_name, last_name, membership_level FROM students WHERE id = ?",
            (student_id,),
        ).fetchone()

        if not student:
            return jsonify({"success": False, "message": "Student not found"})

        # Update membership
        conn.execute(
            "UPDATE students SET membership_level = ?, updated_at = ? WHERE id = ?",
            (new_membership, toronto_now().strftime("%Y-%m-%d %H:%M:%S"), student_id),
        )
        conn.commit()
        conn.close()

        log_action(
            "membership_changed",
            f"Student: {student['first_name']} {student['last_name']}, "
            f"From: {student['membership_level']}, To: {new_membership}",
        )

        return jsonify({"success": True, "message": "Membership updated successfully"})

    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


# ===================================
# SENIOR MANAGER ROUTES
# ===================================


@app.route("/senior/dashboard")
@login_required
@senior_admin_required
def senior_dashboard():
    """Senior manager dashboard with improved data handling"""
    conn = get_db()

    try:
        # Get comprehensive dashboard stats - these are correct values
        stats = get_dashboard_stats()
        print(f"DEBUG: Stats from get_dashboard_stats: {stats}")

        # System usage analytics (last 30 days) with better error handling
        usage_analytics_raw = conn.execute(
            """
            SELECT 
                date(created_at) as date,
                COUNT(*) as cancellations,
                COUNT(DISTINCT student_id) as active_students
            FROM cancellations
            WHERE created_at >= date('now', '-30 days')
            GROUP BY date(created_at)
            ORDER BY date DESC
            LIMIT 30
        """
        ).fetchall()

        # Convert to list of dicts for template
        usage_analytics = []
        for row in usage_analytics_raw:
            usage_analytics.append(
                {
                    "date": row["date"],
                    "cancellations": row["cancellations"] or 0,
                    "active_students": row["active_students"] or 0,
                }
            )

        # Recent system activity with better formatting
        recent_activity_raw = conn.execute(
            """
            SELECT 
                l.action,
                l.details,
                l.created_at,
                l.user_type,
                CASE 
                    WHEN l.user_type = 'client' AND s.first_name IS NOT NULL 
                        THEN s.first_name || ' ' || s.last_name
                    WHEN l.user_type = 'admin' AND a.first_name IS NOT NULL 
                        THEN a.first_name || ' ' || a.last_name
                    ELSE 'System User'
                END as user_name
            FROM system_logs l
            LEFT JOIN students s ON l.user_id = s.id AND l.user_type = 'client'
            LEFT JOIN admin_users a ON l.user_id = a.id AND l.user_type = 'admin'
            WHERE l.created_at >= date('now', '-7 days')
            ORDER BY l.created_at DESC
            LIMIT 15
        """
        ).fetchall()

        # Convert to list of dicts
        recent_activity = []
        for row in recent_activity_raw:
            recent_activity.append(
                {
                    "action": row["action"],
                    "details": row["details"] or "",
                    "created_at": row["created_at"],
                    "user_type": row["user_type"] or "unknown",
                    "user_name": row["user_name"] or "Unknown User",
                }
            )

        # Recent trends for insights
        monthly_trends = conn.execute(
            """
            SELECT 
                strftime('%Y-%m', created_at) as month,
                COUNT(*) as total_cancellations,
                SUM(CASE WHEN charged = 0 THEN 1 ELSE 0 END) as free_cancellations,
                SUM(CASE WHEN charged = 1 THEN 1 ELSE 0 END) as charged_cancellations
            FROM cancellations
            WHERE created_at >= date('now', '-6 months')
            GROUP BY strftime('%Y-%m', created_at)
            ORDER BY month DESC
        """
        ).fetchall()

        conn.close()

        # Make sure all needed properties exist in stats for template compatibility
        final_stats = {
            # Core stats that work correctly
            "active_students": stats.get("active_students", 0),
            "month_cancellations": stats.get("month_cancellations", 0),
            "pending_reviews": stats.get("pending_reviews", 0),
            "active_tiers": stats.get("active_tiers", 7),
            # Additional properties for completeness
            "today_cancellations": stats.get("today_cancellations", 0),
            "today_vs_yesterday": stats.get("today_vs_yesterday", "No data"),
            "free_cancellations": stats.get("free_cancellations", 0),
            "charged_cancellations": stats.get("charged_cancellations", 0),
            "excluded_cancellations": stats.get("excluded_cancellations", 0),
            "new_students_month": stats.get("new_students_month", 0),
            "total_students": stats.get(
                "total_students", stats.get("active_students", 0)
            ),
            "monthly_cancellations": stats.get(
                "monthly_cancellations", stats.get("month_cancellations", 0)
            ),
            "system_alerts": stats.get("system_alerts", 0),
        }

        print(f"DEBUG: Final stats being sent to template: {final_stats}")

        log_action("senior_dashboard_viewed", "Senior manager accessed dashboard")

        return render_template(
            "senior_dashboard.html",
            stats=final_stats,
            usage_analytics=usage_analytics,
            recent_activity=recent_activity,
            monthly_trends=[dict(row) for row in monthly_trends],
        )

    except Exception as e:
        conn.close()
        print(f"ERROR in senior_dashboard: {str(e)}")
        import traceback

        traceback.print_exc()

        # Log error
        log_action(
            "senior_dashboard_error", f"Error loading senior dashboard: {str(e)}"
        )

        # Fallback with minimal data
        fallback_stats = {
            "active_students": 0,
            "month_cancellations": 0,
            "pending_reviews": 0,
            "total_students": 0,
            "active_tiers": 7,
            "monthly_cancellations": 0,
            "today_cancellations": 0,
            "today_vs_yesterday": "Error",
            "free_cancellations": 0,
            "charged_cancellations": 0,
            "excluded_cancellations": 0,
            "new_students_month": 0,
            "system_alerts": 0,
        }

        return render_template(
            "senior_dashboard.html",
            stats=fallback_stats,
            usage_analytics=[],
            recent_activity=[],
            monthly_trends=[],
            error_message=f"Dashboard data unavailable: {str(e)}",
        )


@app.route("/senior/generate-report", methods=["GET"])
@login_required
@senior_admin_required
def generate_system_report():
    """Generate comprehensive system report"""
    try:
        conn = get_db()

        # Collect comprehensive system data
        report_data = {
            "timestamp": datetime.now().isoformat(),
            "students": conn.execute(
                "SELECT COUNT(*) as count FROM students"
            ).fetchone()["count"],
            "cancellations": conn.execute(
                "SELECT COUNT(*) as count FROM cancellations"
            ).fetchone()["count"],
            "this_month": conn.execute(
                "SELECT COUNT(*) as count FROM cancellations WHERE created_at >= date('now', 'start of month')"
            ).fetchone()["count"],
            "pending": conn.execute(
                "SELECT COUNT(*) as count FROM cancellations WHERE status = 'pending'"
            ).fetchone()["count"],
        }

        conn.close()

        # Create simple HTML report
        html_report = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>System Report - {report_data['timestamp'][:10]}</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 40px; }}
                .header {{ border-bottom: 2px solid #333; padding-bottom: 20px; }}
                .metric {{ margin: 20px 0; padding: 15px; background: #f5f5f5; border-radius: 5px; }}
                .metric h3 {{ margin-top: 0; }}
            </style>
        </head>
        <body>
            <div class="header">
                <h1>Riverside Equestrian System Report</h1>
                <p>Generated: {report_data['timestamp']}</p>
            </div>
            
            <div class="metric">
                <h3>Total Students</h3>
                <p>{report_data['students']}</p>
            </div>
            
            <div class="metric">
                <h3>Total Cancellations</h3>
                <p>{report_data['cancellations']}</p>
            </div>
            
            <div class="metric">
                <h3>This Month's Cancellations</h3>
                <p>{report_data['this_month']}</p>
            </div>
            
            <div class="metric">
                <h3>Pending Reviews</h3>
                <p>{report_data['pending']}</p>
            </div>
        </body>
        </html>
        """

        log_action(
            "system_report_generated", f"Generated at {report_data['timestamp']}"
        )

        return html_report

    except Exception as e:
        log_action("system_report_error", f"Failed to generate report: {str(e)}")
        return (
            f"<html><body><h1>Error generating report</h1><p>{str(e)}</p></body></html>"
        )


@app.route("/senior/backup", methods=["POST"])
@login_required
@senior_admin_required
def senior_backup():
    """Create system backup - complete implementation"""
    import shutil
    import os
    import sqlite3
    from datetime import datetime
    import tempfile
    import zipfile

    try:
        # Create timestamp for backup files
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_dir = "backups"

        # Create backups directory if it doesn't exist
        os.makedirs(backup_dir, exist_ok=True)

        # Backup filename
        backup_filename = f"riverside_backup_{timestamp}"
        backup_path = os.path.join(backup_dir, backup_filename)

        # Create a zip file for the complete backup
        zip_path = f"{backup_path}.zip"

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as backup_zip:

            # 1. Backup the database file
            db_backup_path = f"{backup_path}_database.db"

            # Create a proper SQLite backup
            source_conn = sqlite3.connect(app.config["DATABASE"])
            backup_conn = sqlite3.connect(db_backup_path)

            # Use SQLite's backup API for a clean backup
            source_conn.backup(backup_conn)

            # Close connections
            source_conn.close()
            backup_conn.close()

            # Add database to zip
            backup_zip.write(db_backup_path, f"database_{timestamp}.db")

            # 2. Create a SQL dump as well
            sql_dump_path = f"{backup_path}_dump.sql"
            with open(sql_dump_path, "w") as sql_file:
                conn = sqlite3.connect(app.config["DATABASE"])
                for line in conn.iterdump():
                    sql_file.write("%s\n" % line)
                conn.close()

            # Add SQL dump to zip
            backup_zip.write(sql_dump_path, f"database_dump_{timestamp}.sql")

            # 3. Backup configuration and logs if they exist
            config_files = []

            # Add any config files that exist
            if os.path.exists("config.py"):
                config_files.append("config.py")
            if os.path.exists("requirements.txt"):
                config_files.append("requirements.txt")

            for config_file in config_files:
                try:
                    backup_zip.write(config_file, f"config/{config_file}")
                except:
                    pass  # Skip if file doesn't exist or can't be read

            # 4. Create a backup manifest
            manifest_path = f"{backup_path}_manifest.txt"
            with open(manifest_path, "w") as manifest:
                manifest.write(f"Riverside Equestrian System Backup\n")
                manifest.write(f"Created: {datetime.now().isoformat()}\n")
                manifest.write(f"Created by: {session.get('user_email', 'Unknown')}\n")
                manifest.write(f"Database: {app.config['DATABASE']}\n")
                manifest.write(f"Backup includes:\n")
                manifest.write(f"- Complete database backup\n")
                manifest.write(f"- SQL dump\n")
                manifest.write(f"- Configuration files\n")

                # Add some statistics
                conn = get_db()
                try:
                    stats = conn.execute(
                        """
                        SELECT 
                            (SELECT COUNT(*) FROM students) as students,
                            (SELECT COUNT(*) FROM cancellations) as cancellations,
                            (SELECT COUNT(*) FROM admin_users) as admins,
                            (SELECT COUNT(*) FROM membership_tiers) as tiers
                    """
                    ).fetchone()

                    manifest.write(f"\nDatabase Statistics:\n")
                    manifest.write(f"- Students: {stats['students']}\n")
                    manifest.write(f"- Cancellations: {stats['cancellations']}\n")
                    manifest.write(f"- Admin Users: {stats['admins']}\n")
                    manifest.write(f"- Membership Tiers: {stats['tiers']}\n")
                except:
                    manifest.write(f"\nCould not retrieve database statistics\n")
                finally:
                    conn.close()

            # Add manifest to zip
            backup_zip.write(manifest_path, f"backup_manifest_{timestamp}.txt")

        # Clean up temporary files
        temp_files = [db_backup_path, sql_dump_path, manifest_path]
        for temp_file in temp_files:
            try:
                os.remove(temp_file)
            except:
                pass  # Ignore cleanup errors

        # Calculate backup size
        backup_size = os.path.getsize(zip_path)
        backup_size_mb = round(backup_size / (1024 * 1024), 2)

        # Log the successful backup
        log_action(
            "backup_created",
            f"System backup created: {backup_filename}.zip ({backup_size_mb} MB)",
        )

        # Return success response
        return jsonify(
            {
                "success": True,
                "filename": f"{backup_filename}.zip",
                "size_mb": backup_size_mb,
                "path": zip_path,
                "timestamp": timestamp,
                "message": f"Backup created successfully! File size: {backup_size_mb} MB",
            }
        )

    except sqlite3.Error as db_error:
        # Database-specific error
        error_msg = f"Database backup failed: {str(db_error)}"
        log_action("backup_failed", error_msg)
        return jsonify({"success": False, "error": error_msg, "error_type": "database"})

    except OSError as file_error:
        # File system error
        error_msg = f"File system error during backup: {str(file_error)}"
        log_action("backup_failed", error_msg)
        return jsonify(
            {"success": False, "error": error_msg, "error_type": "filesystem"}
        )

    except Exception as e:
        # General error
        error_msg = f"Backup failed: {str(e)}"
        log_action("backup_failed", error_msg)
        return jsonify({"success": False, "error": error_msg, "error_type": "general"})


@app.route("/senior/list-backups", methods=["GET"])
@login_required
@senior_admin_required
def list_backups():
    """List all available backups"""
    try:
        backup_dir = "backups"
        backups = []

        if os.path.exists(backup_dir):
            for filename in os.listdir(backup_dir):
                if filename.endswith(".zip") and filename.startswith(
                    "riverside_backup_"
                ):
                    file_path = os.path.join(backup_dir, filename)
                    file_stats = os.stat(file_path)

                    # Extract timestamp from filename
                    timestamp_str = filename.replace("riverside_backup_", "").replace(
                        ".zip", ""
                    )
                    try:
                        created_date = datetime.strptime(timestamp_str, "%Y%m%d_%H%M%S")
                    except:
                        created_date = datetime.fromtimestamp(file_stats.st_mtime)

                    backups.append(
                        {
                            "filename": filename,
                            "size_mb": round(file_stats.st_size / (1024 * 1024), 2),
                            "created": created_date.isoformat(),
                            "created_readable": created_date.strftime(
                                "%Y-%m-%d %H:%M:%S"
                            ),
                        }
                    )

        # Sort by creation date (newest first)
        backups.sort(key=lambda x: x["created"], reverse=True)

        return jsonify({"success": True, "backups": backups, "count": len(backups)})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/senior/download-backup/<filename>", methods=["GET"])
@login_required
@senior_admin_required
def download_backup(filename):
    """Download a specific backup file"""
    try:
        # Sanitize filename to prevent directory traversal
        safe_filename = os.path.basename(filename)
        if not safe_filename.endswith(".zip") or not safe_filename.startswith(
            "riverside_backup_"
        ):
            return jsonify({"success": False, "error": "Invalid backup file"}), 400

        backup_path = os.path.join("backups", safe_filename)

        if not os.path.exists(backup_path):
            return jsonify({"success": False, "error": "Backup file not found"}), 404

        log_action("backup_downloaded", f"Downloaded backup: {safe_filename}")

        return send_file(
            backup_path,
            as_attachment=True,
            download_name=safe_filename,
            mimetype="application/zip",
        )

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/senior/delete-backup/<filename>", methods=["DELETE"])
@login_required
@senior_admin_required
def delete_backup(filename):
    """Delete a specific backup file"""
    try:
        # Sanitize filename
        safe_filename = os.path.basename(filename)
        if not safe_filename.endswith(".zip") or not safe_filename.startswith(
            "riverside_backup_"
        ):
            return jsonify({"success": False, "error": "Invalid backup file"}), 400

        backup_path = os.path.join("backups", safe_filename)

        if not os.path.exists(backup_path):
            return jsonify({"success": False, "error": "Backup file not found"}), 404

        os.remove(backup_path)
        log_action("backup_deleted", f"Deleted backup: {safe_filename}")

        return jsonify(
            {"success": True, "message": f"Backup {safe_filename} deleted successfully"}
        )

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/senior/settings")
@login_required
@senior_admin_required
def senior_settings():
    """Senior manager settings"""
    conn = get_db()

    # Get all system settings
    settings = {}
    settings_rows = conn.execute("SELECT key, value FROM system_settings").fetchall()
    for setting in settings_rows:
        settings[setting["key"]] = setting["value"]

    conn.close()

    return render_template("senior_settings.html", settings=settings)


@app.route("/senior/tiers")
@login_required
@senior_admin_required
def senior_tiers():
    """Senior manager tier management"""
    conn = get_db()

    # Get all tiers with student counts
    tiers = conn.execute(
        """
        SELECT mt.*, COUNT(s.id) as student_count
        FROM membership_tiers mt
        LEFT JOIN students s ON mt.level = s.membership_level
        GROUP BY mt.level
        ORDER BY mt.sort_order
    """
    ).fetchall()

    conn.close()

    return render_template("senior_tiers.html", tiers=tiers)


# ===================================
# TEMPLATE MANAGEMENT API ROUTES
# ===================================


@app.route("/senior/api/template-variables", methods=["GET"])
@login_required
@senior_admin_required
def get_template_variables_api():
    """Get available template variables with categories"""
    try:
        # Sample variables for demonstration
        sample_variables = {
            "client_name": "John Smith",
            "client_first_name": "John",
            "client_last_name": "Smith",
            "client_email": "john@example.com",
            "client_phone": "604-123-4567",
            "membership_tier": "Silver",
            "parent_name": "Jane Smith",
            "lesson_date": "March 15, 2024",
            "lesson_time": "3:00 PM",
            "cancellation_status": "Free cancellation",
            "company_name": "Riverside Equestrian",
            "contact_email": "managers@riversideequestrian.ca",
            "current_date": "March 10, 2024",
            "policy_url": "https://www.riversideequestrian.ca/cancellations",
        }

        # Organize variables by category
        variable_categories = {
            "Student Information": [
                "client_name",
                "client_first_name",
                "client_last_name",
                "client_email",
                "client_phone",
                "membership_tier",
                "parent_name",
            ],
            "Lesson Details": ["lesson_date", "lesson_time", "cancellation_status"],
            "Company Information": [
                "company_name",
                "contact_email",
                "policy_url",
                "current_date",
            ],
        }

        return jsonify(
            {
                "success": True,
                "variables": sample_variables,
                "categories": variable_categories,
            }
        )

    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


@app.route("/senior/api/template/<template_id>", methods=["GET"])
@login_required
@senior_admin_required
def get_template_for_edit(template_id):
    """Get template data for editing"""
    try:
        conn = get_db()
        template = conn.execute(
            "SELECT * FROM email_templates WHERE id = ?", (template_id,)
        ).fetchone()
        conn.close()

        if not template:
            return jsonify({"success": False, "message": "Template not found"})

        template_dict = dict(template)
        return jsonify({"success": True, "template": template_dict})

    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


@app.route("/senior/api/template/save", methods=["POST"])
@login_required
@senior_admin_required
def save_template_enhanced():
    """Save or update email template"""
    try:
        data = request.json

        # Validate required fields
        required_fields = ["id", "name", "subject", "body", "type"]
        for field in required_fields:
            if not data.get(field):
                return jsonify(
                    {"success": False, "message": f"Missing required field: {field}"}
                )

        # Clean and prepare data
        template_data = {
            "id": data["id"].strip().replace(" ", "_").lower(),
            "name": data["name"].strip(),
            "subject": data["subject"].strip(),
            "body": data["body"],
            "type": data.get("type", "client"),
            "active": bool(data.get("active", True)),
            "auto_send": bool(data.get("auto_send", True)),
            "priority": data.get("priority", "normal"),
            "delay_minutes": int(data.get("delay_minutes", 0)),
            "include_attachment": bool(data.get("include_attachment", False)),
        }

        # Extract variables used in template
        import re

        variables_used = []
        variable_pattern = r"\{\{(\w+)\}\}"

        # Find variables in subject
        subject_vars = re.findall(variable_pattern, template_data["subject"])
        variables_used.extend(subject_vars)

        # Find variables in body
        body_vars = re.findall(variable_pattern, template_data["body"])
        variables_used.extend(body_vars)

        template_data["variables_used"] = ",".join(set(variables_used))

        conn = get_db()

        # Check if template exists
        existing = conn.execute(
            "SELECT id FROM email_templates WHERE id = ?", (template_data["id"],)
        ).fetchone()

        if existing:
            # Update existing template
            conn.execute(
                """
                UPDATE email_templates 
                SET name = ?, subject = ?, body = ?, type = ?, active = ?, 
                    auto_send = ?, priority = ?, delay_minutes = ?, 
                    include_attachment = ?, variables_used = ?, updated_at = ?
                WHERE id = ?
                """,
                (
                    template_data["name"],
                    template_data["subject"],
                    template_data["body"],
                    template_data["type"],
                    template_data["active"],
                    template_data["auto_send"],
                    template_data["priority"],
                    template_data["delay_minutes"],
                    template_data["include_attachment"],
                    template_data["variables_used"],
                    toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                    template_data["id"],
                ),
            )
            action = "updated"
        else:
            # Insert new template
            conn.execute(
                """
                INSERT INTO email_templates 
                (id, name, subject, body, type, active, auto_send, priority, 
                 delay_minutes, include_attachment, variables_used, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    template_data["id"],
                    template_data["name"],
                    template_data["subject"],
                    template_data["body"],
                    template_data["type"],
                    template_data["active"],
                    template_data["auto_send"],
                    template_data["priority"],
                    template_data["delay_minutes"],
                    template_data["include_attachment"],
                    template_data["variables_used"],
                    toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                    toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                ),
            )
            action = "created"

        conn.commit()
        conn.close()

        return jsonify(
            {
                "success": True,
                "message": f"Template {action} successfully",
                "template_id": template_data["id"],
            }
        )

    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


@app.route("/senior/api/template/<template_id>/delete", methods=["DELETE"])
@login_required
@senior_admin_required
def delete_template_enhanced(template_id):
    """Delete email template"""
    try:
        conn = get_db()

        # Check if template exists
        template = conn.execute(
            "SELECT name FROM email_templates WHERE id = ?", (template_id,)
        ).fetchone()

        if not template:
            return jsonify({"success": False, "message": "Template not found"})

        # Delete template
        conn.execute("DELETE FROM email_templates WHERE id = ?", (template_id,))
        conn.commit()
        conn.close()

        return jsonify({"success": True, "message": "Template deleted successfully"})

    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


@app.route("/senior/api/template/<template_id>/toggle", methods=["POST"])
@login_required
@senior_admin_required
def toggle_template_enhanced(template_id):
    """Toggle template active status"""
    try:
        conn = get_db()

        # Get current status
        template = conn.execute(
            "SELECT active, name FROM email_templates WHERE id = ?", (template_id,)
        ).fetchone()

        if not template:
            return jsonify({"success": False, "message": "Template not found"})

        new_status = not bool(template["active"])

        # Update status
        conn.execute(
            "UPDATE email_templates SET active = ?, updated_at = ? WHERE id = ?",
            (new_status, toronto_now().strftime("%Y-%m-%d %H:%M:%S"), template_id),
        )
        conn.commit()
        conn.close()

        status_text = "activated" if new_status else "deactivated"
        return jsonify(
            {
                "success": True,
                "message": f"Template '{template['name']}' {status_text}",
                "active": new_status,
            }
        )

    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


@app.route("/senior/api/template/<template_id>/duplicate", methods=["POST"])
@login_required
@senior_admin_required
def duplicate_template_enhanced(template_id):
    """Duplicate an email template"""
    try:
        conn = get_db()
        template = conn.execute(
            "SELECT * FROM email_templates WHERE id = ?", (template_id,)
        ).fetchone()

        if not template:
            return jsonify({"success": False, "message": "Template not found"})

        template_dict = dict(template)

        # Create new template data
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        new_id = f"{template_id}_copy_{timestamp}"

        # Insert duplicated template
        conn.execute(
            """
            INSERT INTO email_templates 
            (id, name, subject, body, type, active, auto_send, priority, 
             delay_minutes, include_attachment, variables_used, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                new_id,
                f"Copy of {template_dict['name']}",
                template_dict["subject"],
                template_dict["body"],
                template_dict["type"],
                False,  # Start inactive
                template_dict.get("auto_send", True),
                template_dict.get("priority", "normal"),
                template_dict.get("delay_minutes", 0),
                template_dict.get("include_attachment", False),
                template_dict.get("variables_used", ""),
                toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )
        conn.commit()
        conn.close()

        log_action("template_duplicated", f"Duplicated {template_id} to {new_id}")

        return jsonify(
            {
                "success": True,
                "message": "Template duplicated successfully",
                "new_id": new_id,
            }
        )

    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


@app.route("/senior/api/template/<template_id>/test", methods=["POST"])
@login_required
@senior_admin_required
def send_template_test_enhanced(template_id):
    """Send test email for a template"""
    try:
        data = request.json
        test_email = data.get("test_email")

        if not test_email:
            return jsonify({"success": False, "message": "Test email address required"})

        # Validate email format
        import re

        email_pattern = r"^[^@]+@[^@]+\.[^@]+$"
        if not re.match(email_pattern, test_email):
            return jsonify({"success": False, "message": "Invalid email format"})

        # Get template
        conn = get_db()
        template = conn.execute(
            "SELECT * FROM email_templates WHERE id = ?", (template_id,)
        ).fetchone()
        conn.close()

        if not template:
            return jsonify({"success": False, "message": "Template not found"})

        template_dict = dict(template)

        # Create sample data for testing
        sample_student = {
            "first_name": "Sarah",
            "last_name": "Johnson",
            "parent_first": "Michael",
            "parent_last": "Johnson",
            "email": test_email,
            "phone": "604-123-4567",
            "membership_level": "Silver",
            "id": 999,
        }

        sample_cancellation = {
            "lesson_date": "2024-03-15",
            "lesson_time": "15:00:00",
            "charged": False,
            "sequential_lessons": None,
            "reschedule_requested": False,
            "reschedule_preferences": "",
            "error_report": "",
            "manager_notes": "",
            "id": 999,
        }

        # Generate template variables
        variables = get_template_variables(
            sample_student,
            sample_cancellation,
            {
                "status_message": "This is a test email with sample data.",
                "test_mode": "true",
            },
        )

        # Process template
        body, subject = process_template_variables(
            template_dict["body"], template_dict["subject"], variables
        )

        # Add test disclaimer
        test_subject = f"[TEST] {subject}"
        test_body = f"""
        <div style="background-color: #fff3cd; border: 1px solid #ffeaa7; padding: 15px; margin-bottom: 20px; border-radius: 5px;">
            <h4 style="color: #856404; margin: 0 0 10px 0;">⚠️ TEST EMAIL</h4>
            <p style="margin: 0; color: #856404;">
                This is a test email with sample data from template: <strong>{template_dict['name']}</strong><br>
                Template ID: {template_id}<br>
                Sent from: {email_config.from_email}<br>
                Test sent at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
            </p>
        </div>
        {body}
        <hr style="margin: 30px 0;">
        <p style="font-size: 12px; color: #666;">
            <strong>Template Variables Used:</strong><br>
            {', '.join([f'{{{{{k}}}}}' for k in variables.keys() if f'{{{{{k}}}}}' in template_dict['body'] or f'{{{{{k}}}}}' in template_dict['subject']])}
        </p>
        """

        # Send test email
        result = send_email(
            test_email,
            test_subject,
            test_body,
            "test",
            template_id=f"test_{template_id}",
        )

        if result["success"]:
            log_action("test_email_sent", f"Template: {template_id}, To: {test_email}")

        return jsonify(result)

    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


@app.route("/senior/api/template/<template_id>/preview", methods=["GET"])
@login_required
@senior_admin_required
def preview_template_enhanced(template_id):
    """Get template preview with sample data"""
    try:
        conn = get_db()
        template = conn.execute(
            "SELECT * FROM email_templates WHERE id = ?", (template_id,)
        ).fetchone()
        conn.close()

        if not template:
            return jsonify({"success": False, "message": "Template not found"})

        template_dict = dict(template)

        # Create sample data
        sample_variables = {
            "client_name": "Sarah Johnson",
            "client_first_name": "Sarah",
            "lesson_date": "March 15, 2024",
            "lesson_time": "3:00 PM",
            "membership_tier": "Silver",
            "cancellation_status": "Free cancellation",
            "company_name": "Riverside Equestrian",
            "current_date": datetime.now().strftime("%B %d, %Y"),
        }

        # Simple variable replacement
        preview_body = template_dict["body"]
        preview_subject = template_dict["subject"]

        for key, value in sample_variables.items():
            preview_body = preview_body.replace(f"{{{{{key}}}}}", str(value))
            preview_subject = preview_subject.replace(f"{{{{{key}}}}}", str(value))

        return jsonify(
            {
                "success": True,
                "preview": {
                    "subject": preview_subject,
                    "body": preview_body,
                    "template_name": template_dict["name"],
                    "template_type": template_dict["type"],
                },
            }
        )

    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


@app.route("/senior/api/templates/export", methods=["GET"])
@login_required
@senior_admin_required
def export_templates_enhanced():
    """Export all email templates to JSON"""
    try:
        conn = get_db()
        templates = conn.execute(
            "SELECT * FROM email_templates ORDER BY name"
        ).fetchall()
        conn.close()

        # Convert to list of dictionaries
        template_list = [dict(template) for template in templates]

        # Create export data
        export_data = {
            "export_info": {
                "exported_at": datetime.now().isoformat(),
                "exported_by": session.get("user_email", "Unknown"),
                "system": "Riverside Equestrian Cancellation System",
                "version": "1.0",
            },
            "templates": template_list,
        }

        # Create JSON response
        import json

        response_data = json.dumps(export_data, indent=2, default=str)

        response = app.response_class(
            response_data,
            mimetype="application/json",
            headers={
                "Content-Disposition": f"attachment; filename=riverside_email_templates_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            },
        )

        return response

    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


@app.route("/senior/api/templates/import", methods=["POST"])
@login_required
@senior_admin_required
def import_templates_enhanced():
    """Import email templates from JSON"""
    try:
        data = request.json

        # Handle both old and new format
        if "templates" in data:
            templates = data["templates"]
        else:
            templates = data

        if not templates:
            return jsonify(
                {"success": False, "message": "No templates found in import data"}
            )

        imported_count = 0
        updated_count = 0
        errors = []

        conn = get_db()

        for template_data in templates:
            try:
                # Validate required fields
                if not all(
                    key in template_data for key in ["id", "name", "subject", "body"]
                ):
                    errors.append(
                        f"Template missing required fields: {template_data.get('id', 'unknown')}"
                    )
                    continue

                # Check if template exists
                existing = conn.execute(
                    "SELECT id FROM email_templates WHERE id = ?",
                    (template_data["id"],),
                ).fetchone()

                if existing:
                    # Update existing
                    conn.execute(
                        """
                        UPDATE email_templates 
                        SET name = ?, subject = ?, body = ?, type = ?, active = ?, 
                            auto_send = ?, priority = ?, delay_minutes = ?, 
                            include_attachment = ?, variables_used = ?, updated_at = ?
                        WHERE id = ?
                        """,
                        (
                            template_data["name"],
                            template_data["subject"],
                            template_data["body"],
                            template_data.get("type", "client"),
                            template_data.get("active", True),
                            template_data.get("auto_send", True),
                            template_data.get("priority", "normal"),
                            template_data.get("delay_minutes", 0),
                            template_data.get("include_attachment", False),
                            template_data.get("variables_used", ""),
                            toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                            template_data["id"],
                        ),
                    )
                    updated_count += 1
                else:
                    # Insert new
                    conn.execute(
                        """
                        INSERT INTO email_templates 
                        (id, name, subject, body, type, active, auto_send, priority, 
                         delay_minutes, include_attachment, variables_used, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            template_data["id"],
                            template_data["name"],
                            template_data["subject"],
                            template_data["body"],
                            template_data.get("type", "client"),
                            template_data.get("active", True),
                            template_data.get("auto_send", True),
                            template_data.get("priority", "normal"),
                            template_data.get("delay_minutes", 0),
                            template_data.get("include_attachment", False),
                            template_data.get("variables_used", ""),
                            toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                            toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
                        ),
                    )
                    imported_count += 1

            except Exception as e:
                errors.append(
                    f"Template {template_data.get('id', 'unknown')}: {str(e)}"
                )

        conn.commit()
        conn.close()

        return jsonify(
            {
                "success": True,
                "message": f"Import completed: {imported_count} new, {updated_count} updated",
                "imported": imported_count,
                "updated": updated_count,
                "errors": errors[:10],  # Limit error list
            }
        )

    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


# ===================================
# UPDATE EXISTING SENIOR TEMPLATES ROUTE
# ===================================


@app.route("/senior/templates")
@login_required
@senior_admin_required
def senior_templates():
    """Senior manager email templates page"""
    try:
        conn = get_db()

        # Get all email templates
        templates = conn.execute(
            """
            SELECT id, name, subject, body, type, active, auto_send, priority, 
                   delay_minutes, include_attachment, variables_used, created_at, updated_at
            FROM email_templates 
            ORDER BY active DESC, type, name
        """
        ).fetchall()

        conn.close()

        # Convert to list of dicts for template
        template_list = []
        for template in templates:
            template_dict = dict(template)

            # Parse dates properly
            for date_field in ["created_at", "updated_at"]:
                if template_dict.get(date_field):
                    try:
                        template_dict[date_field] = datetime.strptime(
                            template_dict[date_field], "%Y-%m-%d %H:%M:%S"
                        )
                    except (ValueError, TypeError):
                        template_dict[date_field] = datetime.now()

            template_list.append(template_dict)

        return render_template("senior_templates.html", templates=template_list)

    except Exception as e:
        print(f"Error in senior_templates: {e}")
        return render_template("senior_templates.html", templates=[], error=str(e))


# Senior Manager API endpoints
@app.route("/senior/save-settings", methods=["POST"])
@login_required
@senior_admin_required
def save_settings():
    """Save system settings"""
    data = request.json

    conn = get_db()

    try:
        # Update each setting
        for form_name, form_data in data.items():
            for key, value in form_data.items():
                conn.execute(
                    "INSERT OR REPLACE INTO system_settings (key, value, updated_at) VALUES (?, ?, ?)",
                    (key, str(value), datetime.now().isoformat()),
                )

        conn.commit()
        conn.close()

        log_action("settings_updated", f"Updated settings: {list(data.keys())}")
        return jsonify({"success": True})
    except Exception as e:
        conn.close()
        return jsonify({"success": False, "error": str(e)})


@app.route("/senior/save-tiers", methods=["POST"])
@login_required
@senior_admin_required
def save_tiers():
    """Save membership tiers"""
    data = request.json
    tiers = data.get("tiers", [])

    conn = get_db()

    try:
        for tier in tiers:
            conn.execute(
                """
                UPDATE membership_tiers 
                SET free_notices = ?, deadline_hours = ?, deadline_display = ?, active = ?
                WHERE level = ?
            """,
                (
                    tier["limit"],
                    18 if tier["deadline"] == "6pm_previous_day" else 2,
                    tier["deadline"].replace("_", " "),
                    1 if tier["status"] == "active" else 0,
                    tier["id"],
                ),
            )

        conn.commit()
        conn.close()

        log_action("tiers_updated", f"Updated {len(tiers)} tiers")
        return jsonify({"success": True})
    except Exception as e:
        conn.close()
        return jsonify({"success": False, "error": str(e)})


@app.route("/senior/save-template", methods=["POST"])
@login_required
@senior_admin_required
def save_template():
    """Save email template"""
    data = request.json

    conn = get_db()

    try:
        conn.execute(
            """
            INSERT OR REPLACE INTO email_templates 
            (id, name, subject, body, type, active, auto_send, priority, delay_minutes, include_attachment, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
            (
                data["id"],
                data["id"].replace("_", " ").title(),
                data["subject"],
                data["body"],
                data["type"],
                data["active"],
                data["autoSend"],
                data["priority"],
                data["delay"],
                data["includeAttachment"],
                datetime.now().isoformat(),
            ),
        )

        conn.commit()
        conn.close()

        log_action("template_updated", f"Template: {data['id']}")
        return jsonify({"success": True})
    except Exception as e:
        conn.close()
        return jsonify({"success": False, "error": str(e)})


@app.route("/senior/migrate-students", methods=["POST"])
@login_required
@senior_admin_required
def migrate_students():
    """Migrate students between tiers"""
    data = request.json
    from_tier = data.get("from_tier")
    to_tier = data.get("to_tier")

    conn = get_db()

    try:
        # Count students to migrate
        count = conn.execute(
            "SELECT COUNT(*) as count FROM students WHERE membership_level = ?",
            (from_tier,),
        ).fetchone()["count"]

        # Perform migration
        conn.execute(
            "UPDATE students SET membership_level = ? WHERE membership_level = ?",
            (to_tier, from_tier),
        )

        conn.commit()
        conn.close()

        log_action(
            "students_migrated",
            f"Migrated {count} students from {from_tier} to {to_tier}",
        )
        return jsonify({"success": True, "count": count})
    except Exception as e:
        conn.close()
        return jsonify({"success": False, "error": str(e)})


@app.route("/senior/create-backup", methods=["POST"])
@login_required
@senior_admin_required
def create_backup():
    """Create system backup"""
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"backup_{timestamp}.sql"

        # In a real implementation, you would create an actual database backup
        # For this demo, we'll just return a success message

        log_action("backup_created", f"Backup file: {backup_filename}")
        return jsonify({"success": True, "filename": backup_filename})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/test-email")
def test_email_route():
    result = test_email_configuration()
    if result["success"]:
        # Send a test email to yourself
        test_result = send_email(
            "stav@riversideequestrian.ca",
            "Test Email from Riverside System",
            "<h1>Success!</h1><p>Your email system is working correctly!</p>",
            template_id="test",
        )
        return (
            f"Email config: {result['message']}<br>Test email: {test_result['message']}"
        )
    else:
        return f"Email config failed: {result['message']}"


@app.route("/test-timezone")
def test_timezone_route():
    """Debug route to test timezone functionality"""
    import time

    toronto_time = toronto_now()
    system_time = datetime.now()
    utc_time = datetime.utcnow()

    html_output = f"""
    <h2>Timezone Debug Information</h2>
    <p><strong>Current Toronto Time:</strong> {toronto_time}</p>
    <p><strong>Formatted Toronto Time:</strong> {format_datetime_for_display(toronto_time)}</p>
    <p><strong>System Time:</strong> {system_time}</p>
    <p><strong>UTC Time:</strong> {utc_time}</p>
    <p><strong>Toronto Timezone:</strong> {toronto_time.tzname()}</p>
    <p><strong>UTC Offset:</strong> {toronto_time.strftime('%z')}</p>
    <p><strong>Is DST Active:</strong> {'Yes (EDT)' if toronto_time.dst().total_seconds() > 0 else 'No (EST)'}</p>
    <p><strong>System TZ:</strong> {time.tzname}</p>
    <hr>
    <p><em>The Toronto time should be 5 hours behind UTC in winter (EST) or 4 hours behind in summer (EDT).</em></p>
    """

    return html_output


# ===================================
# API ENDPOINTS FOR AJAX CALLS
# ===================================


@app.route("/api/preview-cancellation", methods=["POST"])
@login_required
def preview_cancellation():
    """Preview cancellation status before submission"""
    if session["user_role"] != "client":
        return jsonify({"error": "Unauthorized"}), 403

    data = request.json
    lesson_date = data.get("lesson_date")
    lesson_time = data.get("lesson_time")

    if not lesson_date or not lesson_time:
        return jsonify({"error": "Missing date or time"}), 400

    try:
        lesson_datetime = datetime.strptime(
            f"{lesson_date} {lesson_time}", "%Y-%m-%d %H:%M"
        )
    except ValueError:
        return jsonify({"error": "Invalid date or time format"}), 400

    # Get student info
    conn = get_db()
    student = conn.execute(
        "SELECT * FROM students WHERE id = ?", (session["user_id"],)
    ).fetchone()
    conn.close()

    if not student:
        return jsonify({"error": "Student not found"}), 404

    # Calculate status
    will_charge, reason = will_be_charged(student, lesson_datetime)
    status = calculate_cancellation_status(student)

    return jsonify(
        {"will_be_charged": will_charge, "reason": reason, "current_status": status}
    )


# ===================================
# TEMPLATE CONTEXT PROCESSORS
# ===================================


@app.context_processor
def inject_now():
    """Inject current datetime into all templates"""
    return {"now": toronto_now()}


# Replace your current inject_moment context processor with this:


@app.context_processor
def inject_moment():
    """Inject moment-like function for templates"""

    def moment_func(date_str):
        """Format relative time like moment.js"""
        if not date_str:
            return "Unknown"

        try:
            # Parse the date string
            if isinstance(date_str, str):
                date_obj = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
            else:
                date_obj = date_str

            now = datetime.now()
            time_diff = now - date_obj

            total_seconds = int(time_diff.total_seconds())

            if total_seconds < 60:
                return "Just now"
            elif total_seconds < 3600:
                minutes = total_seconds // 60
                return f"{minutes} minute{'s' if minutes != 1 else ''} ago"
            elif total_seconds < 86400:
                hours = total_seconds // 3600
                return f"{hours} hour{'s' if hours != 1 else ''} ago"
            elif total_seconds < 2592000:  # 30 days
                days = total_seconds // 86400
                return f"{days} day{'s' if days != 1 else ''} ago"
            else:
                return date_obj.strftime("%Y-%m-%d")

        except (ValueError, TypeError):
            return str(date_str)

    return {"moment": moment_func}


# ===================================
# STATIC FILE FIXES
# ===================================


@app.route("/static/js/main.js")
def serve_main_js():
    """Serve main.js file"""
    js_content = """
// Main JavaScript file for Riverside Equestrian
console.log('Riverside Equestrian Cancellation System loaded');

// Global utilities
window.RiversideUtils = {
    formatDate: function(date) {
        return new Date(date).toLocaleDateString();
    },
    
    formatTime: function(time) {
        return new Date('1970-01-01T' + time).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    },
    
    showAlert: function(message, type = 'info') {
        const alertDiv = document.createElement('div');
        alertDiv.className = `alert alert-${type} alert-dismissible fade show position-fixed`;
        alertDiv.style.cssText = 'top: 20px; right: 20px; z-index: 9999; min-width: 300px;';
        alertDiv.innerHTML = `
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        `;
        document.body.appendChild(alertDiv);
        
        setTimeout(() => {
            if (alertDiv.parentNode) {
                alertDiv.remove();
            }
        }, 5000);
    }
};

// Initialize tooltips if Bootstrap is available
document.addEventListener('DOMContentLoaded', function() {
    if (typeof bootstrap !== 'undefined' && bootstrap.Tooltip) {
        var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
        var tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {
            return new bootstrap.Tooltip(tooltipTriggerEl);
        });
    }
});
"""
    return app.response_class(js_content, mimetype="application/javascript")


@app.route("/favicon.ico")
def serve_favicon():
    """Serve favicon or return 204 No Content"""
    return "", 204


@app.route("/manager/export/cancellations")
@login_required
@admin_required
def export_cancellations_excel():
    """Export cancellations with current filters to Excel"""
    try:
        conn = get_db()

        # Get filter parameters from the request - same as manager_cancellations route
        filter_status = request.args.get("status", "")
        search = request.args.get("search", "")
        date_range = request.args.get("date_range", "month")
        membership = request.args.get("membership", "")
        sort_by = request.args.get("sort", "submit_date")
        student_id = request.args.get("student")
        date_from = request.args.get("date_from")
        date_to = request.args.get("date_to")

        # Build the same query as manager_cancellations route
        where_clauses = []
        params = []

        # Add student filter
        if student_id:
            where_clauses.append("c.student_id = ?")
            params.append(student_id)

        # Status filters
        if filter_status == "free":
            where_clauses.append("c.charged = 0 AND c.status = 'approved'")
        elif filter_status == "charged":
            where_clauses.append("c.charged = 1")
        elif filter_status == "excluded":
            where_clauses.append("c.excluded = 1")
        elif filter_status == "note":
            where_clauses.append(
                "c.cancellation_note IS NOT NULL AND c.cancellation_note != ''"
            )
        elif filter_status == "deadline_passed":
            where_clauses.append("c.deadline_passed = 1")
        elif filter_status == "override":
            where_clauses.append("c.is_override = 1")

        # Search filter
        if search:
            where_clauses.append(
                "(s.first_name LIKE ? OR s.last_name LIKE ? OR s.email LIKE ?)"
            )
            search_param = f"%{search}%"
            params.extend([search_param, search_param, search_param])

        # Date range filter
        if date_range == "today":
            where_clauses.append("DATE(c.created_at) = DATE('now')")
        elif date_range == "yesterday":
            where_clauses.append("DATE(c.created_at) = DATE('now', '-1 day')")
        elif date_range == "7days":
            where_clauses.append("c.created_at >= DATE('now', '-7 days')")
        elif date_range == "month":
            where_clauses.append("c.created_at >= DATE('now', '-30 days')")
        elif date_range == "all":
            pass  # No date filter
        elif date_range == "custom" and date_from and date_to:
            where_clauses.append("DATE(c.created_at) BETWEEN ? AND ?")
            params.extend([date_from, date_to])

        # Membership filter
        if membership:
            where_clauses.append("s.membership_level = ?")
            params.append(membership)

        where_sql = " AND ".join(where_clauses) if where_clauses else "1=1"

        # Sort order
        if sort_by == "submit_date":
            order_by = "c.created_at DESC"
        elif sort_by == "lesson_date":
            order_by = "c.lesson_date DESC"
        elif sort_by == "student":
            order_by = "s.last_name, s.first_name"
        elif sort_by == "status":
            order_by = "c.status, c.charged, c.excluded"
        else:
            order_by = "c.created_at DESC"

        # Get cancellations data with the same query as the page
        cancellations = conn.execute(
            f"""
            SELECT 
                c.id,
                s.first_name,
                s.last_name,
                s.parent_first,
                s.parent_last,
                s.email,
                s.membership_level,
                c.lesson_date,
                c.lesson_time,
                c.created_at,
                c.charged,
                c.excluded,
                c.status,
                c.deadline_passed,
                c.is_override,
                c.manager_notes,
                c.cancellation_note
            FROM cancellations c
            JOIN students s ON c.student_id = s.id
            WHERE {where_sql}
            ORDER BY {order_by}
        """,
            params,
        ).fetchall()

        conn.close()

        # Create Excel workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Cancellations Export"

        # Headers
        headers = [
            "ID",
            "Student First Name",
            "Student Last Name",
            "Parent First Name",
            "Parent Last Name",
            "Email",
            "Membership Level",
            "Lesson Date",
            "Lesson Time",
            "Submitted Date",
            "Status",
            "Cancellation Note",
            "Manager Notes",
        ]

        # Add headers
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(
                start_color="E6F3FF", end_color="E6F3FF", fill_type="solid"
            )

        # Add data
        for row, cancellation in enumerate(cancellations, 2):
            ws.cell(row=row, column=1, value=cancellation["id"])
            ws.cell(row=row, column=2, value=cancellation["first_name"])
            ws.cell(row=row, column=3, value=cancellation["last_name"])
            ws.cell(row=row, column=4, value=cancellation["parent_first"] or "")
            ws.cell(row=row, column=5, value=cancellation["parent_last"] or "")
            ws.cell(row=row, column=6, value=cancellation["email"])
            ws.cell(row=row, column=7, value=cancellation["membership_level"])
            ws.cell(row=row, column=8, value=cancellation["lesson_date"])
            ws.cell(row=row, column=9, value=cancellation["lesson_time"])

            # Format submitted date
            submitted_date = datetime.fromisoformat(
                cancellation["created_at"]
            ).strftime("%Y-%m-%d %H:%M")
            ws.cell(row=row, column=10, value=submitted_date)

            # Determine status
            if cancellation["excluded"]:
                status = "Excluded"
            elif cancellation["charged"]:
                status = "Charged"
            elif cancellation["status"] == "approved":
                status = "Free"
            elif cancellation["deadline_passed"]:
                status = "Deadline Passed"
            elif cancellation["is_override"]:
                status = "Override"
            else:
                status = "Pending"

            ws.cell(row=row, column=11, value=status)
            ws.cell(row=row, column=12, value=cancellation["cancellation_note"] or "")
            ws.cell(row=row, column=13, value=cancellation["manager_notes"] or "")

        # Auto-adjust column widths
        for column in ws.columns:
            max_length = max(len(str(cell.value or "")) for cell in column)
            ws.column_dimensions[column[0].column_letter].width = min(
                max_length + 2, 50
            )

        # Save to temporary file
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        wb.save(temp_file.name)
        temp_file.close()

        # Generate filename with current filters
        filter_parts = []
        if filter_status:
            filter_parts.append(filter_status)
        if date_range and date_range != "all":
            filter_parts.append(date_range)
        if search:
            filter_parts.append("search")

        filename_suffix = "_".join(filter_parts) if filter_parts else "all"
        filename = (
            f'cancellations_{filename_suffix}_{datetime.now().strftime("%Y%m%d")}.xlsx'
        )

        log_action(
            "export_cancellations", f"Exported filtered cancellations: {filename}"
        )

        return send_file(
            temp_file.name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename,
        )

    except Exception as e:
        log_action("export_error", f"Error exporting filtered cancellations: {str(e)}")
        return jsonify({"success": False, "message": f"Export failed: {str(e)}"}), 500


@app.route("/manager/export/cancellations/monthly")
@login_required
@admin_required
def export_cancellations_monthly_excel():
    """Export monthly cancellation report to Excel"""
    try:
        conn = get_db()

        # Get current month and previous months data
        current_date = datetime.now()
        start_date = current_date.replace(day=1) - relativedelta(months=5)

        # Get cancellations data
        cancellations = conn.execute(
            """
            SELECT 
                c.id,
                s.first_name,
                s.last_name,
                s.parent_first,
                s.parent_last,
                s.email,
                s.membership_level,
                c.lesson_date,
                c.lesson_time,
                c.created_at,
                c.charged,
                c.excluded,
                c.manager_notes,
                c.cancellation_note
            FROM cancellations c
            JOIN students s ON c.student_id = s.id
            WHERE c.created_at >= ?
            ORDER BY c.created_at DESC
        """,
            (start_date.strftime("%Y-%m-%d"),),
        ).fetchall()

        conn.close()

        # Create Excel workbook
        wb = Workbook()

        # Create summary sheet
        ws_summary = wb.active
        ws_summary.title = "Monthly Summary"

        # Summary headers
        headers = [
            "Month",
            "Total Cancellations",
            "Free",
            "Charged",
            "Excluded",
            "Bronze Members",
            "Silver Members",
            "Gold Members",
        ]

        for col, header in enumerate(headers, 1):
            cell = ws_summary.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(
                start_color="366092", end_color="366092", fill_type="solid"
            )
            cell.alignment = Alignment(horizontal="center")

        # Calculate monthly summaries
        monthly_data = {}
        for cancellation in cancellations:
            month_key = datetime.fromisoformat(cancellation["created_at"]).strftime(
                "%Y-%m"
            )
            if month_key not in monthly_data:
                monthly_data[month_key] = {
                    "total": 0,
                    "free": 0,
                    "charged": 0,
                    "excluded": 0,
                    "bronze": 0,
                    "silver": 0,
                    "gold": 0,
                }

            monthly_data[month_key]["total"] += 1
            if cancellation["excluded"]:
                monthly_data[month_key]["excluded"] += 1
            elif cancellation["charged"]:
                monthly_data[month_key]["charged"] += 1
            else:
                monthly_data[month_key]["free"] += 1

            # Count by membership
            membership = cancellation["membership_level"].lower()
            if membership == "bronze":
                monthly_data[month_key]["bronze"] += 1
            elif membership == "silver":
                monthly_data[month_key]["silver"] += 1
            elif membership == "gold":
                monthly_data[month_key]["gold"] += 1

        # Add summary data
        row = 2
        for month in sorted(monthly_data.keys(), reverse=True):
            data = monthly_data[month]
            month_name = datetime.strptime(month, "%Y-%m").strftime("%B %Y")

            ws_summary.cell(row=row, column=1, value=month_name)
            ws_summary.cell(row=row, column=2, value=data["total"])
            ws_summary.cell(row=row, column=3, value=data["free"])
            ws_summary.cell(row=row, column=4, value=data["charged"])
            ws_summary.cell(row=row, column=5, value=data["excluded"])
            ws_summary.cell(row=row, column=6, value=data["bronze"])
            ws_summary.cell(row=row, column=7, value=data["silver"])
            ws_summary.cell(row=row, column=8, value=data["gold"])
            row += 1

        # Auto-adjust column widths
        for column in ws_summary.columns:
            max_length = max(len(str(cell.value or "")) for cell in column)
            ws_summary.column_dimensions[column[0].column_letter].width = max_length + 2

        # Create detailed sheet
        ws_detail = wb.create_sheet("Detailed Cancellations")

        # Detailed headers
        detail_headers = [
            "ID",
            "Student Name",
            "Parent Name",
            "Email",
            "Membership",
            "Lesson Date",
            "Lesson Time",
            "Submitted Date",
            "Status",
            "Cancellation Note",
            "Manager Notes",
        ]

        for col, header in enumerate(detail_headers, 1):
            cell = ws_detail.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(
                start_color="366092", end_color="366092", fill_type="solid"
            )
            cell.alignment = Alignment(horizontal="center")

        # Add detailed data
        for row, cancellation in enumerate(cancellations, 2):
            ws_detail.cell(row=row, column=1, value=cancellation["id"])

            student_name = f"{cancellation['first_name']} {cancellation['last_name']}"
            ws_detail.cell(row=row, column=2, value=student_name)

            parent_name = ""
            if cancellation["parent_first"] or cancellation["parent_last"]:
                parent_name = f"{cancellation['parent_first'] or ''} {cancellation['parent_last'] or ''}".strip()
            ws_detail.cell(row=row, column=3, value=parent_name)

            ws_detail.cell(row=row, column=4, value=cancellation["email"])
            ws_detail.cell(row=row, column=5, value=cancellation["membership_level"])
            ws_detail.cell(row=row, column=6, value=cancellation["lesson_date"])
            ws_detail.cell(row=row, column=7, value=cancellation["lesson_time"])

            submitted_date = datetime.fromisoformat(
                cancellation["created_at"]
            ).strftime("%Y-%m-%d %H:%M")
            ws_detail.cell(row=row, column=8, value=submitted_date)

            if cancellation["excluded"]:
                status = "Excluded"
            elif cancellation["charged"]:
                status = "Charged"
            else:
                status = "Free"
            ws_detail.cell(row=row, column=9, value=status)

            ws_detail.cell(
                row=row, column=10, value=cancellation["cancellation_note"] or ""
            )
            ws_detail.cell(
                row=row, column=11, value=cancellation["manager_notes"] or ""
            )

        # Auto-adjust column widths for detailed sheet
        for column in ws_detail.columns:
            max_length = max(len(str(cell.value or "")) for cell in column)
            ws_detail.column_dimensions[column[0].column_letter].width = min(
                max_length + 2, 50
            )

        # Save to temporary file
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        wb.save(temp_file.name)
        temp_file.close()

        log_action("export_cancellations", f"Exported cancellations report")

        return send_file(
            temp_file.name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f'cancellations_report_{datetime.now().strftime("%Y%m%d")}.xlsx',
        )

    except Exception as e:
        log_action("export_error", f"Error exporting cancellations: {str(e)}")
        return jsonify({"success": False, "message": f"Export failed: {str(e)}"}), 500


@app.route("/manager/export/students")
@login_required
@admin_required
def export_students_excel():
    """Export students to Excel with enhanced formatting"""
    try:
        conn = get_db()

        # Get students data with cancellation statistics
        students = conn.execute(
            """
            SELECT 
                s.*,
                COUNT(c.id) as total_cancellations,
                SUM(CASE WHEN c.charged = 0 AND c.excluded = 0 THEN 1 ELSE 0 END) as free_cancellations,
                SUM(CASE WHEN c.charged = 1 THEN 1 ELSE 0 END) as charged_cancellations,
                SUM(CASE WHEN c.excluded = 1 THEN 1 ELSE 0 END) as excluded_cancellations,
                MAX(c.created_at) as last_cancellation_date
            FROM students s
            LEFT JOIN cancellations c ON s.id = c.student_id
            GROUP BY s.id
            ORDER BY s.last_name, s.first_name
        """
        ).fetchall()

        conn.close()

        # Create Excel workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Students Directory"

        # Headers
        headers = [
            "ID",
            "First Name",
            "Last Name",
            "Parent First",
            "Parent Last",
            "Email",
            "Phone",
            "Membership Level",
            "Created Date",
            "Total Cancellations",
            "Free",
            "Charged",
            "Excluded",
            "Last Cancellation",
        ]

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(
                start_color="366092", end_color="366092", fill_type="solid"
            )
            cell.alignment = Alignment(horizontal="center")

        # Add student data
        for row, student in enumerate(students, 2):
            ws.cell(row=row, column=1, value=student["id"])
            ws.cell(row=row, column=2, value=student["first_name"])
            ws.cell(row=row, column=3, value=student["last_name"])
            ws.cell(row=row, column=4, value=student["parent_first"] or "")
            ws.cell(row=row, column=5, value=student["parent_last"] or "")
            ws.cell(row=row, column=6, value=student["email"])
            ws.cell(row=row, column=7, value=student["phone"] or "")
            ws.cell(row=row, column=8, value=student["membership_level"])

            created_date = datetime.fromisoformat(student["created_at"]).strftime(
                "%Y-%m-%d"
            )
            ws.cell(row=row, column=9, value=created_date)

            ws.cell(row=row, column=10, value=student["total_cancellations"])
            ws.cell(row=row, column=11, value=student["free_cancellations"])
            ws.cell(row=row, column=12, value=student["charged_cancellations"])
            ws.cell(row=row, column=13, value=student["excluded_cancellations"])

            last_cancellation = ""
            if student["last_cancellation_date"]:
                last_cancellation = datetime.fromisoformat(
                    student["last_cancellation_date"]
                ).strftime("%Y-%m-%d")
            ws.cell(row=row, column=14, value=last_cancellation)

        # Auto-adjust column widths
        for column in ws.columns:
            max_length = max(len(str(cell.value or "")) for cell in column)
            ws.column_dimensions[column[0].column_letter].width = min(
                max_length + 2, 50
            )

        # Save to temporary file
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        wb.save(temp_file.name)
        temp_file.close()

        log_action("export_students", f"Exported {len(students)} students")

        return send_file(
            temp_file.name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f'students_directory_{datetime.now().strftime("%Y%m%d")}.xlsx',
        )

    except Exception as e:
        log_action("export_error", f"Error exporting students: {str(e)}")
        return jsonify({"success": False, "message": f"Export failed: {str(e)}"}), 500


@app.route("/manager/export/charged-cancellations")
@login_required
@admin_required
def export_charged_cancellations():
    """Export charged cancellations report to Excel"""
    try:
        conn = get_db()

        # Get charged cancellations data
        charged_cancellations = conn.execute(
            """
            SELECT 
                c.id,
                s.first_name,
                s.last_name,
                s.parent_first,
                s.parent_last,
                s.email,
                s.phone,
                s.membership_level,
                c.lesson_date,
                c.lesson_time,
                c.created_at,
                c.cancellation_note,
                c.manager_notes
            FROM cancellations c
            JOIN students s ON c.student_id = s.id
            WHERE c.charged = 1
            ORDER BY c.created_at DESC
        """
        ).fetchall()

        conn.close()

        # Create Excel workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Charged Cancellations"

        # Headers
        headers = [
            "Cancellation ID",
            "Student Name",
            "Parent Name",
            "Email",
            "Phone",
            "Membership Level",
            "Lesson Date",
            "Lesson Time",
            "Submitted Date",
            "Cancellation Note",
            "Manager Notes",
        ]

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(
                start_color="F59E0B", end_color="F59E0B", fill_type="solid"
            )
            cell.alignment = Alignment(horizontal="center")

        # Add data
        for row, cancellation in enumerate(charged_cancellations, 2):
            ws.cell(row=row, column=1, value=cancellation["id"])

            student_name = f"{cancellation['first_name']} {cancellation['last_name']}"
            ws.cell(row=row, column=2, value=student_name)

            parent_name = ""
            if cancellation["parent_first"] or cancellation["parent_last"]:
                parent_name = f"{cancellation['parent_first'] or ''} {cancellation['parent_last'] or ''}".strip()
            ws.cell(row=row, column=3, value=parent_name)

            ws.cell(row=row, column=4, value=cancellation["email"])
            ws.cell(row=row, column=5, value=cancellation["phone"] or "")
            ws.cell(row=row, column=6, value=cancellation["membership_level"])
            ws.cell(row=row, column=7, value=cancellation["lesson_date"])
            ws.cell(row=row, column=8, value=cancellation["lesson_time"])

            submitted_date = datetime.fromisoformat(
                cancellation["created_at"]
            ).strftime("%Y-%m-%d %H:%M")
            ws.cell(row=row, column=9, value=submitted_date)

            ws.cell(row=row, column=10, value=cancellation["cancellation_note"] or "")
            ws.cell(row=row, column=11, value=cancellation["manager_notes"] or "")

        # Auto-adjust column widths
        for column in ws.columns:
            max_length = max(len(str(cell.value or "")) for cell in column)
            ws.column_dimensions[column[0].column_letter].width = min(
                max_length + 2, 50
            )

        # Add summary at top
        ws.insert_rows(1, 2)
        ws.cell(
            row=1,
            column=1,
            value=f"Charged Cancellations Report - Generated {toronto_now().strftime('%Y-%m-%d %H:%M')}",
        )
        ws.cell(row=1, column=1).font = Font(bold=True, size=14)
        ws.cell(
            row=2,
            column=1,
            value=f"Total Charged Cancellations: {len(charged_cancellations)}",
        )
        ws.cell(row=2, column=1).font = Font(bold=True)

        # Save to temporary file
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        wb.save(temp_file.name)
        temp_file.close()

        log_action(
            "export_charged",
            f"Exported {len(charged_cancellations)} charged cancellations",
        )

        return send_file(
            temp_file.name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f'charged_cancellations_{datetime.now().strftime("%Y%m%d")}.xlsx',
        )

    except Exception as e:
        log_action("export_error", f"Error exporting charged cancellations: {str(e)}")
        return jsonify({"success": False, "message": f"Export failed: {str(e)}"}), 500


@app.route("/manager/export/analytics")
@login_required
@admin_required
def export_full_analytics():
    """Export comprehensive analytics report to Excel with multiple sheets"""
    try:
        conn = get_db()

        # Create Excel workbook with multiple sheets
        wb = Workbook()

        # 1. Executive Summary Sheet
        ws_summary = wb.active
        ws_summary.title = "Executive Summary"

        # Get overall statistics
        current_month = datetime.now().replace(day=1)
        last_month = current_month - relativedelta(months=1)

        stats = {}

        # Current month stats
        stats["current_month"] = conn.execute(
            """
            SELECT 
                COUNT(*) as total,
                SUM(CASE WHEN charged = 0 AND excluded = 0 THEN 1 ELSE 0 END) as free,
                SUM(CASE WHEN charged = 1 THEN 1 ELSE 0 END) as charged,
                SUM(CASE WHEN excluded = 1 THEN 1 ELSE 0 END) as excluded
            FROM cancellations 
            WHERE created_at >= ? AND created_at < ?
        """,
            (
                current_month.strftime("%Y-%m-%d"),
                (current_month + relativedelta(months=1)).strftime("%Y-%m-%d"),
            ),
        ).fetchone()

        # Previous month stats
        stats["last_month"] = conn.execute(
            """
            SELECT 
                COUNT(*) as total,
                SUM(CASE WHEN charged = 0 AND excluded = 0 THEN 1 ELSE 0 END) as free,
                SUM(CASE WHEN charged = 1 THEN 1 ELSE 0 END) as charged,
                SUM(CASE WHEN excluded = 1 THEN 1 ELSE 0 END) as excluded
            FROM cancellations 
            WHERE created_at >= ? AND created_at < ?
        """,
            (last_month.strftime("%Y-%m-%d"), current_month.strftime("%Y-%m-%d")),
        ).fetchone()

        # Student counts by membership
        membership_stats = conn.execute(
            """
            SELECT membership_level, COUNT(*) as count
            FROM students
            GROUP BY membership_level
            ORDER BY count DESC
        """
        ).fetchall()

        # Executive Summary content
        summary_data = [
            ["Riverside Equestrian - Analytics Report", "", ""],
            [f"Generated: {toronto_now().strftime('%Y-%m-%d %H:%M')}", "", ""],
            ["", "", ""],
            ["MONTHLY COMPARISON", "", ""],
            ["Metric", "Current Month", "Previous Month"],
            [
                "Total Cancellations",
                stats["current_month"]["total"],
                stats["last_month"]["total"],
            ],
            [
                "Free Cancellations",
                stats["current_month"]["free"],
                stats["last_month"]["free"],
            ],
            [
                "Charged Cancellations",
                stats["current_month"]["charged"],
                stats["last_month"]["charged"],
            ],
            [
                "Excluded Cancellations",
                stats["current_month"]["excluded"],
                stats["last_month"]["excluded"],
            ],
            ["", "", ""],
            ["MEMBERSHIP DISTRIBUTION", "", ""],
            ["Membership Level", "Student Count", "Percentage"],
        ]

        total_students = sum(m["count"] for m in membership_stats)
        for membership in membership_stats:
            percentage = f"{(membership['count'] / total_students * 100):.1f}%"
            summary_data.append(
                [membership["membership_level"], membership["count"], percentage]
            )

        # Write summary data
        for row_idx, row_data in enumerate(summary_data, 1):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws_summary.cell(row=row_idx, column=col_idx, value=value)
                if row_idx in [1, 4, 11]:  # Header rows
                    cell.font = Font(bold=True, size=12)
                elif row_idx == 5 or row_idx == 12:  # Sub-header rows
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(
                        start_color="E5E7EB", end_color="E5E7EB", fill_type="solid"
                    )

        # 2. Monthly Trends Sheet
        ws_trends = wb.create_sheet("Monthly Trends")

        # Get 12 months of data
        start_date = datetime.now().replace(day=1) - relativedelta(months=11)
        monthly_trends = conn.execute(
            """
            SELECT 
                strftime('%Y-%m', created_at) as month,
                COUNT(*) as total,
                SUM(CASE WHEN charged = 0 AND excluded = 0 THEN 1 ELSE 0 END) as free,
                SUM(CASE WHEN charged = 1 THEN 1 ELSE 0 END) as charged,
                SUM(CASE WHEN excluded = 1 THEN 1 ELSE 0 END) as excluded
            FROM cancellations 
            WHERE created_at >= ?
            GROUP BY strftime('%Y-%m', created_at)
            ORDER BY month
        """,
            (start_date.strftime("%Y-%m-%d"),),
        ).fetchall()

        # Write trends data
        trends_headers = ["Month", "Total", "Free", "Charged", "Excluded"]
        for col, header in enumerate(trends_headers, 1):
            cell = ws_trends.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(
                start_color="366092", end_color="366092", fill_type="solid"
            )
            cell.font = Font(bold=True, color="FFFFFF")

        for row, trend in enumerate(monthly_trends, 2):
            month_name = datetime.strptime(trend["month"], "%Y-%m").strftime("%B %Y")
            ws_trends.cell(row=row, column=1, value=month_name)
            ws_trends.cell(row=row, column=2, value=trend["total"])
            ws_trends.cell(row=row, column=3, value=trend["free"])
            ws_trends.cell(row=row, column=4, value=trend["charged"])
            ws_trends.cell(row=row, column=5, value=trend["excluded"])

        # 3. Student Analysis Sheet
        ws_students = wb.create_sheet("Student Analysis")

        # Get top students by cancellation frequency
        top_students = conn.execute(
            """
            SELECT 
                s.first_name,
                s.last_name,
                s.membership_level,
                COUNT(c.id) as total_cancellations,
                SUM(CASE WHEN c.charged = 1 THEN 1 ELSE 0 END) as charged_count,
                MAX(c.created_at) as last_cancellation
            FROM students s
            JOIN cancellations c ON s.id = c.student_id
            GROUP BY s.id
            HAVING COUNT(c.id) > 0
            ORDER BY total_cancellations DESC
            LIMIT 50
        """
        ).fetchall()

        # Write student analysis
        student_headers = [
            "Student Name",
            "Membership",
            "Total Cancellations",
            "Charged",
            "Last Cancellation",
        ]
        for col, header in enumerate(student_headers, 1):
            cell = ws_students.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(
                start_color="10B981", end_color="10B981", fill_type="solid"
            )
            cell.font = Font(bold=True, color="FFFFFF")

        for row, student in enumerate(top_students, 2):
            student_name = f"{student['first_name']} {student['last_name']}"
            ws_students.cell(row=row, column=1, value=student_name)
            ws_students.cell(row=row, column=2, value=student["membership_level"])
            ws_students.cell(row=row, column=3, value=student["total_cancellations"])
            ws_students.cell(row=row, column=4, value=student["charged_count"])

            last_date = ""
            if student["last_cancellation"]:
                last_date = datetime.fromisoformat(
                    student["last_cancellation"]
                ).strftime("%Y-%m-%d")
            ws_students.cell(row=row, column=5, value=last_date)

        # Auto-adjust column widths for all sheets
        for ws in wb.worksheets:
            for column in ws.columns:
                max_length = max(len(str(cell.value or "")) for cell in column)
                ws.column_dimensions[column[0].column_letter].width = min(
                    max_length + 2, 50
                )

        conn.close()

        # Save to temporary file
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        wb.save(temp_file.name)
        temp_file.close()

        log_action("export_analytics", "Exported full analytics report")

        return send_file(
            temp_file.name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f'full_analytics_report_{datetime.now().strftime("%Y%m%d")}.xlsx',
        )

    except Exception as e:
        log_action("export_error", f"Error exporting analytics: {str(e)}")
        return jsonify({"success": False, "message": f"Export failed: {str(e)}"}), 500


# Add route to get cancellation details for management card
@app.route("/manager/api/cancellation/<int:cancellation_id>/card", methods=["GET"])
@login_required
@admin_required
def get_cancellation_card(cancellation_id):
    """Get cancellation data formatted for management card view"""
    try:
        conn = get_db()

        # Get cancellation with student and policy info
        cancellation_data = conn.execute(
            """
            SELECT 
                c.*,
                s.first_name,
                s.last_name,
                s.parent_first,
                s.parent_last,
                s.email,
                s.phone,
                s.membership_level,
                s.created_at as student_created_at
            FROM cancellations c
            JOIN students s ON c.student_id = s.id
            WHERE c.id = ?
        """,
            (cancellation_id,),
        ).fetchone()

        if not cancellation_data:
            return jsonify({"success": False, "message": "Cancellation not found"}), 404

        # Get student's current month cancellation count
        current_month = datetime.now().replace(day=1)
        month_cancellations = conn.execute(
            """
            SELECT COUNT(*) as count
            FROM cancellations
            WHERE student_id = ? AND created_at >= ?
        """,
            (cancellation_data["student_id"], current_month.strftime("%Y-%m-%d")),
        ).fetchone()

        conn.close()

        # Format response
        response_data = {
            "success": True,
            "cancellation": dict(cancellation_data),
            "month_count": month_cancellations["count"],
        }

        return jsonify(response_data)

    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


# ===================================
# IMPROVED ERROR HANDLERS
# ===================================


@app.errorhandler(404)
def not_found_error(error):
    """Handle 404 errors with safe template"""
    return (
        render_template_string(
            """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Page Not Found - Riverside Equestrian</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
    <div class="container mt-5">
        <div class="row justify-content-center">
            <div class="col-md-8 text-center">
                <div class="card shadow">
                    <div class="card-body p-5">
                        <h1 class="display-1 text-primary">404</h1>
                        <h2 class="mb-3">Page Not Found</h2>
                        <p class="lead mb-4">The page you're looking for doesn't exist.</p>
                        <p class="text-muted mb-4">
                            Error occurred at: {{ now.strftime('%Y-%m-%d %H:%M:%S') }}
                        </p>
                        <div class="d-grid gap-2 d-md-block">
                            <a href="{{ url_for('login') }}" class="btn btn-primary">Go to Login</a>
                            <a href="javascript:history.back()" class="btn btn-outline-secondary">Go Back</a>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
    """,
            now=datetime.now(),
        ),
        404,
    )


@app.errorhandler(403)
def forbidden_error(error):
    """Handle 403 errors with safe template"""
    return (
        render_template_string(
            """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Access Forbidden - Riverside Equestrian</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
    <div class="container mt-5">
        <div class="row justify-content-center">
            <div class="col-md-8 text-center">
                <div class="card shadow">
                    <div class="card-body p-5">
                        <h1 class="display-1 text-warning">403</h1>
                        <h2 class="mb-3">Access Forbidden</h2>
                        <p class="lead mb-4">You don't have permission to access this resource.</p>
                        <p class="text-muted mb-4">
                            Error occurred at: {{ now.strftime('%Y-%m-%d %H:%M:%S') }}
                        </p>
                        <div class="d-grid gap-2 d-md-block">
                            <a href="{{ url_for('dashboard') }}" class="btn btn-primary">Go to Dashboard</a>
                            <a href="{{ url_for('logout') }}" class="btn btn-outline-secondary">Logout</a>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
    """,
            now=datetime.now(),
        ),
        403,
    )


@app.errorhandler(500)
def internal_error(error):
    """Handle 500 errors with safe template"""
    return (
        render_template_string(
            """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Server Error - Riverside Equestrian</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
    <div class="container mt-5">
        <div class="row justify-content-center">
            <div class="col-md-8 text-center">
                <div class="card shadow">
                    <div class="card-body p-5">
                        <h1 class="display-1 text-danger">500</h1>
                        <h2 class="mb-3">Server Error</h2>
                        <p class="lead mb-4">An internal server error occurred.</p>
                        <p class="text-muted mb-4">
                            Error occurred at: {{ now.strftime('%Y-%m-%d %H:%M:%S') }}
                        </p>
                        <div class="d-grid gap-2 d-md-block">
                            <a href="{{ url_for('login') }}" class="btn btn-primary">Go to Login</a>
                            <a href="javascript:location.reload()" class="btn btn-outline-secondary">Reload Page</a>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
    """,
            now=datetime.now(),
        ),
        500,
    )


# ===================================
# CREATE STATIC DIRECTORIES
# ===================================


def create_static_structure():
    """Create static file structure if it doesn't exist"""
    static_dirs = ["static", "static/css", "static/js", "static/images"]

    for directory in static_dirs:
        os.makedirs(directory, exist_ok=True)

    # Create basic main.css if it doesn't exist
    main_css_path = "static/css/main.css"
    if not os.path.exists(main_css_path):
        with open(main_css_path, "w") as f:
            f.write(
                """
/* Main CSS for Riverside Equestrian Cancellation System */
:root {
    --primary-color: #4e73df;
    --secondary-color: #858796;
    --success-color: #1cc88a;
    --warning-color: #f6c23e;
    --danger-color: #e74a3b;
    --info-color: #36b9cc;
}

body {
    font-family: 'Nunito', -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Helvetica Neue', Arial, sans-serif;
    background-color: #f8f9fc;
}

.navbar-brand {
    font-weight: 800;
    color: var(--primary-color) !important;
}

.card {
    border: none;
    border-radius: 0.35rem;
    box-shadow: 0 0.15rem 1.75rem 0 rgba(58, 59, 69, 0.15);
}

.card-header {
    background-color: #f8f9fc;
    border-bottom: 1px solid #e3e6f0;
}

.btn-primary {
    background-color: var(--primary-color);
    border-color: var(--primary-color);
}

.btn-primary:hover {
    background-color: #2e59d9;
    border-color: #2e59d9;
}

.alert {
    border: none;
    border-radius: 0.35rem;
}

.form-control:focus {
    border-color: var(--primary-color);
    box-shadow: 0 0 0 0.2rem rgba(78, 115, 223, 0.25);
}

.sidebar {
    background: linear-gradient(180deg, var(--primary-color) 10%, #224abe 100%);
    background-size: cover;
}

.sidebar .nav-link {
    color: rgba(255, 255, 255, 0.8);
}

.sidebar .nav-link:hover {
    color: #fff;
}

.sidebar .nav-link.active {
    color: #fff;
    background-color: rgba(255, 255, 255, 0.1);
    border-radius: 0.35rem;
}

@media (max-width: 768px) {
    .sidebar {
        position: relative !important;
        width: 100% !important;
        height: auto !important;
    }
}

.table th {
    border-top: none;
    font-weight: 800;
    font-size: 0.8rem;
    color: #5a5c69;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.badge {
    font-size: 0.75rem;
    font-weight: 700;
    padding: 0.375rem 0.75rem;
}

.text-truncate {
    max-width: 200px;
}

/* Animation for loading states */
.loading {
    opacity: 0.6;
    pointer-events: none;
}

@keyframes fadeIn {
    from { opacity: 0; transform: translateY(10px); }
    to { opacity: 1; transform: translateY(0); }
}

.fade-in {
    animation: fadeIn 0.3s ease-out;
}

/* Print styles */
@media print {
    .sidebar, .navbar, .btn, .no-print {
        display: none !important;
    }
    
    .container-fluid {
        margin: 0 !important;
        padding: 0 !important;
    }
}
            """
            )

    # Create basic main.js if it doesn't exist
    main_js_path = "static/js/main.js"
    if not os.path.exists(main_js_path):
        with open(main_js_path, "w") as f:
            f.write(
                """
// Main JavaScript for Riverside Equestrian Cancellation System
console.log('Riverside Equestrian system loaded');

// Global utilities
window.RiversideUtils = {
    formatDate: function(date) {
        return new Date(date).toLocaleDateString();
    },
    
    formatTime: function(time) {
        return new Date('1970-01-01T' + time).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    }
};
            """
            )


class EmailConfig:
    """Enhanced Email configuration class with Office 365 support"""

    def __init__(self):
        # Email server configuration
        self.smtp_server = os.getenv("SMTP_SERVER", "smtp.office365.com")
        self.smtp_port = int(os.getenv("SMTP_PORT", "587"))
        self.use_tls = os.getenv("USE_TLS", "True").lower() == "true"
        self.use_ssl = os.getenv("USE_SSL", "False").lower() == "true"
        self.timeout = int(os.getenv("EMAIL_TIMEOUT", "30"))

        # Authentication
        self.smtp_user = os.getenv("SMTP_USER", "")
        self.smtp_password = os.getenv("SMTP_PASSWORD", "")

        # Email addresses
        self.from_email = os.getenv("FROM_EMAIL", "noreply@riversideequestrian.ca")
        self.from_name = os.getenv("FROM_NAME", "Riverside Equestrian")
        self.admin_email = os.getenv("ADMIN_EMAIL", "stav@riversideequestrian.ca")
        self.manager_emails = [
            email.strip()
            for email in os.getenv(
                "MANAGER_EMAIL", "managers@riversideequestrian.ca"
            ).split(",")
        ]

        # Email behavior
        self.send_emails = os.getenv("SEND_EMAILS", "True").lower() == "true"
        self.debug_mode = os.getenv("EMAIL_DEBUG", "False").lower() == "true"
        self.log_emails = True

        # Validate configuration
        self.is_configured = bool(self.smtp_user and self.smtp_password)

    def get_connection(self):
        """Get SMTP connection"""
        try:
            if self.use_ssl:
                # Use SSL connection
                context = ssl.create_default_context()
                server = smtplib.SMTP_SSL(
                    self.smtp_server,
                    self.smtp_port,
                    context=context,
                    timeout=self.timeout,
                )
            else:
                # Use regular SMTP with STARTTLS (Office 365 default)
                server = smtplib.SMTP(
                    self.smtp_server, self.smtp_port, timeout=self.timeout
                )
                if self.use_tls:
                    server.starttls()

            if self.smtp_user and self.smtp_password:
                server.login(self.smtp_user, self.smtp_password)

            return server
        except Exception as e:
            raise Exception(f"Failed to connect to email server: {str(e)}")


email_config = EmailConfig()

# ===================================
# EMAIL SENDING FUNCTIONS
# ===================================


def send_email_async(func):
    """Decorator to send emails asynchronously"""

    @wraps(func)
    def wrapper(*args, **kwargs):
        thread = threading.Thread(target=func, args=args, kwargs=kwargs)
        thread.daemon = True
        thread.start()
        return thread

    return wrapper


def log_email_attempt(to_email, subject, success, error=None, template_id=None):
    """Enhanced email logging"""
    try:
        conn = get_db()
        details = f"To: {to_email}, Subject: {subject}"
        if template_id:
            details += f", Template: {template_id}"
        if error:
            details += f", Error: {error}"

        conn.execute(
            """
            INSERT INTO system_logs (user_id, user_type, action, details, created_at)
            VALUES (?, ?, ?, ?, ?)
        """,
            (
                session.get("user_id"),
                "email_system",
                "email_sent" if success else "email_failed",
                details,
                toronto_now().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )
        conn.commit()
        conn.close()
    except Exception as e:
        if email_config.debug_mode:
            print(f"Failed to log email attempt: {e}")


def send_email(
    to_email, subject, body, template_type="client", attachments=None, template_id=None
):
    """
    Enhanced email sending with Office 365 support

    Args:
        to_email (str): Recipient email address
        subject (str): Email subject
        body (str): HTML email body
        template_type (str): Type of email template
        attachments (list): List of file paths to attach
        template_id (str): Template ID for logging

    Returns:
        dict: Result with success status and message
    """

    if not email_config.send_emails:
        if email_config.debug_mode:
            print(f"Email sending disabled - would send to {to_email}: {subject}")
        return {"success": True, "message": "Email sending disabled (debug mode)"}

    if not email_config.is_configured:
        error_msg = "Email not configured - missing SMTP credentials"
        log_email_attempt(to_email, subject, False, error_msg, template_id)
        return {"success": False, "message": error_msg}

    try:
        # Create message
        msg = MIMEMultipart("alternative")
        msg["From"] = f"{email_config.from_name} <{email_config.from_email}>"
        msg["To"] = to_email
        msg["Subject"] = subject
        msg["Reply-To"] = email_config.from_email

        # Add HTML body
        html_part = MIMEText(body, "html", "utf-8")
        msg.attach(html_part)

        # Add plain text version for better compatibility
        from html import unescape
        import re

        plain_text = re.sub("<[^<]+?>", "", body)
        plain_text = unescape(plain_text)
        text_part = MIMEText(plain_text, "plain", "utf-8")
        msg.attach(text_part)

        # Add attachments if provided
        if attachments:
            for file_path in attachments:
                try:
                    with open(file_path, "rb") as attachment:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(attachment.read())
                        encoders.encode_base64(part)
                        part.add_header(
                            "Content-Disposition",
                            f"attachment; filename= {os.path.basename(file_path)}",
                        )
                        msg.attach(part)
                except Exception as e:
                    if email_config.debug_mode:
                        print(f"Failed to attach file {file_path}: {e}")

        # Send email using connection
        server = email_config.get_connection()
        server.send_message(msg)
        server.quit()

        result = {"success": True, "message": "Email sent successfully"}

        if email_config.debug_mode:
            print(f"✓ Email sent to {to_email}: {subject}")

        log_email_attempt(to_email, subject, True, template_id=template_id)
        return result

    except Exception as e:
        error_msg = str(e)
        result = {"success": False, "message": f"Failed to send email: {error_msg}"}

        if email_config.debug_mode:
            print(f"✗ Email failed to {to_email}: {error_msg}")

        log_email_attempt(to_email, subject, False, error_msg, template_id)
        return result


@send_email_async
def send_email_async_wrapper(
    to_email, subject, body, template_type="client", attachments=None, template_id=None
):
    """Async wrapper for sending emails"""
    return send_email(to_email, subject, body, template_type, attachments, template_id)


# ===================================
# TEMPLATE PROCESSING FUNCTIONS
# ===================================


def get_email_template(template_id):
    """Get email template from database with error handling"""
    try:
        conn = get_db()
        template = conn.execute(
            "SELECT * FROM email_templates WHERE id = ? AND active = 1", (template_id,)
        ).fetchone()
        conn.close()
        return dict(template) if template else None
    except Exception as e:
        if email_config.debug_mode:
            print(f"Error getting template {template_id}: {e}")
        return None


def process_template_variables(template_body, template_subject, variables):
    """
    Enhanced template variable processing with better error handling

    Args:
        template_body (str): Template HTML body
        template_subject (str): Template subject
        variables (dict): Variables to replace

    Returns:
        tuple: (processed_body, processed_subject)
    """
    processed_body = template_body
    processed_subject = template_subject

    for key, value in variables.items():
        placeholder = f"{{{{{key}}}}}"
        processed_body = processed_body.replace(placeholder, str(value))
        processed_subject = processed_subject.replace(placeholder, str(value))

    return processed_body, processed_subject


# ===================================
# TEMPLATE VARIABLE GENERATORS
# ===================================


def get_template_variables(student=None, cancellation=None, extra_vars=None):
    """Generate template variables for your cancellation system - FIXED for sqlite3.Row"""
    variables = {
        "current_date": datetime.now().strftime("%B %d, %Y"),
        "current_time": datetime.now().strftime("%I:%M %p"),
        "company_name": "Riverside Equestrian",
        "contact_email": email_config.admin_email,
        "policy_url": "https://www.riversideequestrian.ca/cancellations",
        "website_url": "https://www.riversideequestrian.ca",
        "system_email": email_config.from_email,
    }

    if student:
        # Convert sqlite3.Row to dict if needed
        if hasattr(student, "keys"):
            student_dict = dict(student)
        else:
            student_dict = student

        variables.update(
            {
                "client_name": f"{student_dict['first_name']} {student_dict['last_name']}",
                "client_first_name": student_dict["first_name"],
                "client_last_name": student_dict["last_name"],
                "client_email": student_dict["email"],
                "client_phone": student_dict.get("phone", ""),
                "membership_tier": student_dict["membership_level"],
                "parent_name": f"{student_dict.get('parent_first', '')} {student_dict.get('parent_last', '')}".strip(),
                "parent_first_name": student_dict.get("parent_first", ""),
                "parent_last_name": student_dict.get("parent_last", ""),
            }
        )

        # Get membership tier info
        tier = get_membership_tier(student_dict["membership_level"])
        if tier:
            tier_dict = dict(tier) if hasattr(tier, "keys") else tier
            variables.update(
                {
                    "allowed_cancellations": str(tier_dict["free_notices"]),
                    "cancellation_deadline": tier_dict["deadline_display"],
                }
            )

        # Get current usage
        status = calculate_cancellation_status(student_dict)
        variables.update(
            {
                "used_cancellations": str(status["used"]),
                "remaining_cancellations": str(status["remaining"]),
            }
        )

    if cancellation:
        # Convert sqlite3.Row to dict if needed
        if hasattr(cancellation, "keys"):
            cancellation_dict = dict(cancellation)
        else:
            cancellation_dict = cancellation

        # Format dates consistently
        lesson_date_str = cancellation_dict.get("lesson_date")
        if isinstance(lesson_date_str, str):
            try:
                lesson_date = datetime.strptime(lesson_date_str, "%Y-%m-%d")
            except ValueError:
                lesson_date = datetime.now()
        else:
            lesson_date = lesson_date_str or datetime.now()

        lesson_time_str = cancellation_dict.get("lesson_time")
        if isinstance(lesson_time_str, str):
            try:
                if len(lesson_time_str.split(":")) == 3:
                    lesson_time = datetime.strptime(lesson_time_str, "%H:%M:%S").time()
                else:
                    lesson_time = datetime.strptime(lesson_time_str, "%H:%M").time()
            except ValueError:
                lesson_time = datetime.now().time()
        else:
            lesson_time = lesson_time_str or datetime.now().time()

        variables.update(
            {
                "lesson_date": lesson_date.strftime("%B %d, %Y"),
                "lesson_time": lesson_time.strftime("%I:%M %p"),
                "cancellation_status": (
                    "Free cancellation"
                    if not cancellation_dict.get("charged")
                    else "Charged cancellation"
                ),
                "will_be_charged": "Yes" if cancellation_dict.get("charged") else "No",
                "charge_reason": get_charge_reason(cancellation_dict),
                "submission_time": datetime.now().strftime("%B %d, %Y at %I:%M %p"),
                "cancellation_id": str(cancellation_dict.get("id", "")),
            }
        )

        # Sequential lessons
        sequential_lessons = cancellation_dict.get("sequential_lessons")
        if sequential_lessons:
            variables["sequential_lessons"] = format_sequential_lessons(
                sequential_lessons
            )
        else:
            variables["sequential_lessons"] = "No additional lessons"

        # Reschedule info
        variables.update(
            {
                "reschedule_requested": (
                    "Yes" if cancellation_dict.get("reschedule_requested") else "No"
                ),
                "reschedule_preferences": cancellation_dict.get(
                    "reschedule_preferences", "None provided"
                ),
                "error_report": cancellation_dict.get("error_report", "None reported"),
            }
        )

    # Add extra variables
    if extra_vars:
        variables.update(extra_vars)

    return variables


def get_charge_reason(cancellation):
    """Get human-readable reason for cancellation charge - FIXED for sqlite3.Row"""
    # Convert sqlite3.Row to dict if needed
    if hasattr(cancellation, "keys"):
        cancellation_dict = dict(cancellation)
    else:
        cancellation_dict = cancellation

    if not cancellation_dict.get("charged"):
        return "Within policy - no charge applied"

    if cancellation_dict.get("manager_notes"):
        return cancellation_dict["manager_notes"]

    # Default reasons based on common scenarios
    return "Late cancellation or monthly limit exceeded"


def format_sequential_lessons(sequential_data):
    """Format sequential lessons for email - FIXED"""
    if not sequential_data:
        return "No additional lessons"

    try:
        if isinstance(sequential_data, str):
            sequential_lessons = eval(sequential_data)
        else:
            sequential_lessons = sequential_data

        if not sequential_lessons:
            return "No additional lessons"

        formatted_lessons = []
        for lesson in sequential_lessons:
            if isinstance(lesson, dict):
                date_str = lesson.get("date", "")
                time_str = lesson.get("time", "")
            else:
                # Handle other formats
                continue

            if date_str and time_str:
                try:
                    lesson_date = datetime.strptime(str(date_str), "%Y-%m-%d").strftime(
                        "%B %d, %Y"
                    )
                    lesson_time = datetime.strptime(str(time_str), "%H:%M").strftime(
                        "%I:%M %p"
                    )
                    formatted_lessons.append(f"{lesson_date} at {lesson_time}")
                except:
                    formatted_lessons.append(f"{date_str} at {time_str}")

        return (
            f"Additional lessons: {', '.join(formatted_lessons)}"
            if formatted_lessons
            else "No additional lessons"
        )

    except Exception as e:
        if email_config.debug_mode:
            print(f"Error formatting sequential lessons: {e}")
        return "No additional lessons"


def safe_dict_convert(row_or_dict):
    """Safely convert sqlite3.Row to dict, or return dict as-is"""
    if row_or_dict is None:
        return {}
    if hasattr(row_or_dict, "keys"):
        return dict(row_or_dict)
    return row_or_dict


# ===================================
# EMAIL SENDING WRAPPER FUNCTIONS (FOR YOUR CANCELLATION SYSTEM)
# ===================================


def debug_email_templates():
    """Debug function to check email template status"""
    print("=" * 60)
    print("DEBUGGING EMAIL TEMPLATES")
    print("=" * 60)

    # Check what templates exist in database
    conn = get_db()
    templates = conn.execute(
        "SELECT id, name, active FROM email_templates ORDER BY id"
    ).fetchall()
    conn.close()

    print(f"Templates in database ({len(templates)} found):")
    for template in templates:
        print(
            f"  - {template['id']}: {template['name']} (Active: {template['active']})"
        )

    print()

    # Test template retrieval
    required_templates = [
        "client_confirmation",
        "cancellation_charged",
        "free_cancellation",
        "manager_notification",
    ]

    for template_id in required_templates:
        template = get_email_template(template_id)
        if template:
            print(f"✅ {template_id}: Found and active")
        else:
            print(f"❌ {template_id}: NOT FOUND or INACTIVE")

    print("\n" + "=" * 60)

    # Test cancellation logic
    print("TESTING CANCELLATION LOGIC")
    print("=" * 60)

    # Mock Bronze student data (like you)
    mock_student = {
        "id": 1,
        "first_name": "Test",
        "last_name": "Student",
        "email": "test@example.com",
        "membership_level": "Bronze",
    }

    # Mock 4th cancellation (should be charged)
    mock_charged_cancellation = {
        "id": 999,
        "lesson_date": "2024-12-15",
        "lesson_time": "14:00:00",
        "charged": True,  # This should trigger charged email
        "manager_notes": "Monthly limit exceeded",
        "cancellation_note": "Family emergency",
    }

    # Mock 1st cancellation (should be free)
    mock_free_cancellation = {
        "id": 998,
        "lesson_date": "2024-12-20",
        "lesson_time": "15:00:00",
        "charged": False,  # This should trigger free email
        "manager_notes": "Within policy",
        "cancellation_note": None,
    }

    print("Test 1: Charged cancellation (Bronze member, 4th cancellation)")
    print(
        f"  - Student: {mock_student['first_name']} {mock_student['last_name']} ({mock_student['membership_level']})"
    )
    print(f"  - Charged status: {mock_charged_cancellation['charged']}")
    print(f"  - Expected template: cancellation_charged")

    # Test the logic
    if mock_charged_cancellation.get("charged"):
        expected_template = "cancellation_charged"
        template = get_email_template("cancellation_charged")
        if template:
            print(f"  ✅ Would use: {expected_template}")
        else:
            print(
                f"  ⚠️  Template '{expected_template}' not found, would fall back to client_confirmation"
            )
    else:
        print(f"  ❌ Logic error: charged=True but treated as free")

    print()
    print("Test 2: Free cancellation (Bronze member, 1st cancellation)")
    print(
        f"  - Student: {mock_student['first_name']} {mock_student['last_name']} ({mock_student['membership_level']})"
    )
    print(f"  - Charged status: {mock_free_cancellation['charged']}")
    print(f"  - Expected template: free_cancellation")

    if not mock_free_cancellation.get("charged"):
        expected_template = "free_cancellation"
        template = get_email_template("free_cancellation")
        if template:
            print(f"  ✅ Would use: {expected_template}")
        else:
            print(
                f"  ⚠️  Template '{expected_template}' not found, would fall back to client_confirmation"
            )
    else:
        print(f"  ❌ Logic error: charged=False but treated as charged")

    print("\n" + "=" * 60)
    return templates


# Add this route to your app.py to run the debug:
@app.route("/debug-email-templates")
@login_required
@senior_admin_required
def debug_email_route():
    """Debug route to check email template status"""
    templates = debug_email_templates()

    output = "<h1>Email Template Debug</h1><pre>"

    # Capture the debug output
    import io
    import sys
    from contextlib import redirect_stdout

    f = io.StringIO()
    with redirect_stdout(f):
        debug_email_templates()
    output += f.getvalue()
    output += "</pre>"

    return output


# MANUAL TEST FUNCTION - Add this too:
def test_email_sending_logic():
    """Test the actual email sending with your current logic"""
    print("\n" + "=" * 60)
    print("TESTING ACTUAL EMAIL SENDING LOGIC")
    print("=" * 60)

    # Your actual student data (Bronze member)
    conn = get_db()
    bronze_student = conn.execute(
        "SELECT * FROM students WHERE membership_level = 'Bronze' LIMIT 1"
    ).fetchone()

    if not bronze_student:
        print("❌ No Bronze student found in database")
        conn.close()
        return

    print(
        f"Testing with: {bronze_student['first_name']} {bronze_student['last_name']} ({bronze_student['membership_level']})"
    )

    # Test charged cancellation
    charged_cancellation = {
        "id": 9999,
        "lesson_date": "2024-12-15",
        "lesson_time": "14:00:00",
        "charged": True,  # Should use cancellation_charged template
        "manager_notes": "Test charged cancellation",
    }

    print("\n1. Testing CHARGED cancellation:")
    print(f"   charged = {charged_cancellation['charged']}")

    # Run your actual function
    try:
        result = send_cancellation_confirmation(bronze_student, charged_cancellation)
        print(f"   Result: {result}")
    except Exception as e:
        print(f"   ❌ Error: {e}")

    # Test free cancellation
    free_cancellation = {
        "id": 9998,
        "lesson_date": "2024-12-20",
        "lesson_time": "15:00:00",
        "charged": False,  # Should use free_cancellation template
        "manager_notes": "Test free cancellation",
    }

    print("\n2. Testing FREE cancellation:")
    print(f"   charged = {free_cancellation['charged']}")

    try:
        result = send_cancellation_confirmation(bronze_student, free_cancellation)
        print(f"   Result: {result}")
    except Exception as e:
        print(f"   ❌ Error: {e}")

    conn.close()


# FIXED EMAIL SENDING FUNCTIONS FOR RIVERSIDE EQUESTRIAN
# Replace your existing functions in app.py with these:


def send_cancellation_confirmation(student, cancellation):
    """Send appropriate cancellation email to client based on charge status - FIXED"""

    # Convert sqlite3.Row to dict if needed
    if hasattr(student, "keys"):
        student_dict = dict(student)
    else:
        student_dict = student

    if hasattr(cancellation, "keys"):
        cancellation_dict = dict(cancellation)
    else:
        cancellation_dict = cancellation

    # CHOOSE THE RIGHT TEMPLATE BASED ON CHARGE STATUS
    if cancellation_dict.get("charged"):
        # Use charged template for paid cancellations
        template_id = "cancellation_charged"
        template = get_email_template("cancellation_charged")
        if not template:
            if email_config.debug_mode:
                print(
                    "Warning: No cancellation_charged template found, falling back to client_confirmation"
                )
            template_id = "client_confirmation"
            template = get_email_template("client_confirmation")
    else:
        # Use free cancellation template for free cancellations
        template_id = "free_cancellation"
        template = get_email_template("free_cancellation")
        if not template:
            if email_config.debug_mode:
                print(
                    "Warning: No free_cancellation template found, falling back to client_confirmation"
                )
            template_id = "client_confirmation"
            template = get_email_template("client_confirmation")

    if not template:
        if email_config.debug_mode:
            print("Error: No email template found at all!")
        return {"success": False, "message": "Template not found"}

    # Generate variables
    variables = get_template_variables(student_dict, cancellation_dict)

    # Add status message based on charge status
    if cancellation_dict.get("charged"):
        variables["status_message"] = (
            f"A charge will be applied to your account. Reason: {get_charge_reason(cancellation_dict)}"
        )
        variables["charge_reason"] = get_charge_reason(cancellation_dict)
    else:
        variables["status_message"] = (
            "This cancellation has been processed at no charge."
        )

    # Process template
    body, subject = process_template_variables(
        template["body"], template["subject"], variables
    )

    # Send email with correct template ID
    result = send_email(
        student_dict["email"],
        subject,
        body,
        "client",
        template_id=template_id,  # This will now be the correct template
    )

    if email_config.debug_mode:
        print(
            f"Sent {template_id} template to {student_dict['email']}, charged={cancellation_dict.get('charged', False)}"
        )

    return result


def send_appropriate_cancellation_email(student, cancellation):
    """
    IMPROVED VERSION: Send the right email based on cancellation status
    This replaces send_cancellation_confirmation with better logic
    """

    # Convert sqlite3.Row to dict if needed
    if hasattr(student, "keys"):
        student_dict = dict(student)
    else:
        student_dict = student

    if hasattr(cancellation, "keys"):
        cancellation_dict = dict(cancellation)
    else:
        cancellation_dict = cancellation

    # Debug info
    is_charged = bool(cancellation_dict.get("charged", False))
    student_name = (
        f"{student_dict.get('first_name', '')} {student_dict.get('last_name', '')}"
    )

    if email_config.debug_mode:
        print(f"DEBUG: Sending email for {student_name}")
        print(f"DEBUG: Cancellation charged status: {is_charged}")
        print(
            f"DEBUG: Will use template: {'cancellation_charged' if is_charged else 'free_cancellation'}"
        )

    # STEP 1: Choose the right template
    if is_charged:
        # This is a charged cancellation - use charged template
        template_id = "cancellation_charged"
        fallback_template_id = "client_confirmation"
        print(
            f"✅ Using CHARGED template for {student_name} (Bronze member, 4th cancellation)"
        )
    else:
        # This is a free cancellation - use free template
        template_id = "free_cancellation"
        fallback_template_id = "client_confirmation"
        print(f"✅ Using FREE template for {student_name}")

    # STEP 2: Get the template
    template = get_email_template(template_id)
    if not template:
        print(f"⚠️  Primary template '{template_id}' not found, using fallback")
        template_id = fallback_template_id
        template = get_email_template(fallback_template_id)

    if not template:
        error_msg = f"❌ No email template found (tried {template_id} and {fallback_template_id})"
        print(error_msg)
        return {"success": False, "message": error_msg}

    # STEP 3: Generate variables for the template
    variables = get_template_variables(student_dict, cancellation_dict)

    # Add charge-specific variables
    if is_charged:
        variables["charge_reason"] = get_charge_reason(cancellation_dict)
        variables["status_message"] = (
            f"A charge will be applied to your account. Reason: {variables['charge_reason']}"
        )
    else:
        variables["status_message"] = (
            "This cancellation has been processed at no charge."
        )

    # STEP 4: Process the template
    try:
        body, subject = process_template_variables(
            template["body"], template["subject"], variables
        )
    except Exception as e:
        error_msg = f"❌ Template processing failed: {str(e)}"
        print(error_msg)
        return {"success": False, "message": error_msg}

    # STEP 5: Send the email
    result = send_email(
        student_dict["email"],
        subject,
        body,
        "client",
        template_id=template_id,
    )

    # Log the result
    if result.get("success"):
        print(f"✅ SUCCESS: Sent '{template_id}' email to {student_dict['email']}")
    else:
        print(
            f"❌ FAILED: Could not send '{template_id}' email to {student_dict['email']}: {result.get('message', 'Unknown error')}"
        )

    return result


def send_manager_notification(student, cancellation):
    """Send new cancellation notification to managers - FIXED for sqlite3.Row"""
    template = get_email_template("manager_notification")
    if not template:
        return {"success": False, "message": "Template not found"}

    # Convert sqlite3.Row to dict if needed
    if hasattr(student, "keys"):
        student_dict = dict(student)
    else:
        student_dict = student

    if hasattr(cancellation, "keys"):
        cancellation_dict = dict(cancellation)
    else:
        cancellation_dict = cancellation

    variables = get_template_variables(
        student_dict,
        cancellation_dict,
        {
            "action_required": "Review cancellation and approve/charge as needed",
            "dashboard_url": f"{request.url_root if 'request' in globals() else 'http://localhost:5000/'}manager/cancellations?student={student_dict['id']}",
        },
    )

    body, subject = process_template_variables(
        template["body"], template["subject"], variables
    )

    # Send to all manager emails
    results = []
    for manager_email in email_config.manager_emails:
        result = send_email(
            manager_email, subject, body, "manager", template_id="manager_notification"
        )
        results.append(result)

    # Return success if any email was sent successfully
    success_count = sum(1 for r in results if r.get("success"))
    return {
        "success": success_count > 0,
        "message": f"Sent to {success_count}/{len(results)} managers",
    }


def send_override_notification_emails(
    student, cancellation, override_action, override_reason, manager_email
):
    """Send emails after manager override - both to client and managers"""

    # Convert data to dicts for consistency
    if hasattr(student, "keys"):
        student_dict = dict(student)
    else:
        student_dict = student

    if hasattr(cancellation, "keys"):
        cancellation_dict = dict(cancellation)
    else:
        cancellation_dict = cancellation

    results = {
        "client_email": {"success": False},
        "manager_notification": {"success": False},
    }

    # Add override-specific information
    override_info = {
        "override_action": override_action,
        "override_reason": override_reason,
        "override_by": manager_email,
        "override_date": datetime.now().strftime("%B %d, %Y at %I:%M %p"),
    }

    # STEP 1: Send updated confirmation to CLIENT
    try:
        print(f"📧 Sending override notification to client: {student_dict['email']}")
        print(
            f"   Action: {override_action}, Charged: {cancellation_dict.get('charged', False)}"
        )

        # Use the standard cancellation confirmation logic (it will pick the right template)
        client_result = send_cancellation_confirmation(student_dict, cancellation_dict)
        results["client_email"] = client_result

        if client_result.get("success"):
            print(f"✅ Client override email sent successfully")
        else:
            print(
                f"❌ Client override email failed: {client_result.get('message', 'Unknown error')}"
            )

    except Exception as e:
        print(f"❌ Client override email error: {str(e)}")
        results["client_email"] = {"success": False, "message": str(e)}

    # STEP 2: Send notification to MANAGERS about the override
    try:
        # Create override notification for managers
        manager_template = get_email_template("manager_override_notification")

        # If no specific override template, use regular manager notification
        if not manager_template:
            manager_template = get_email_template("manager_notification")

        if manager_template:
            # Generate template variables with override info
            variables = get_template_variables(
                student_dict, cancellation_dict, override_info
            )

            # Add manager-specific variables
            variables.update(
                {
                    "manager_action": f"Override: {override_action}",
                    "action_taken": f"{manager_email} {override_action}d this cancellation",
                    "override_summary": f"This cancellation was {override_action}d by {manager_email}",
                    "new_status": (
                        "Free cancellation"
                        if not cancellation_dict.get("charged")
                        else "Charged cancellation"
                    ),
                    "dashboard_url": f"{request.url_root if 'request' in globals() else 'http://localhost:5000/'}manager/cancellations?student={student_dict['id']}",
                }
            )

            # Process template
            body, subject = process_template_variables(
                manager_template["body"],
                f"🔄 Override: {override_action.title()} - {variables['client_name']} - {variables['lesson_date']}",
                variables,
            )

            # Send to all manager emails EXCEPT the one who made the override
            manager_emails = [
                email for email in email_config.manager_emails if email != manager_email
            ]

            if manager_emails:
                manager_results = []
                for mgr_email in manager_emails:
                    result = send_email(
                        mgr_email,
                        subject,
                        body,
                        "manager",
                        template_id="manager_override_notification",
                    )
                    manager_results.append(result)

                success_count = sum(1 for r in manager_results if r.get("success"))
                results["manager_notification"] = {
                    "success": success_count > 0,
                    "message": f"Sent override notification to {success_count}/{len(manager_results)} managers",
                }

                print(
                    f"📧 Manager override notifications sent to {success_count}/{len(manager_results)} managers"
                )
            else:
                results["manager_notification"] = {
                    "success": True,
                    "message": "No other managers to notify (override made by only manager)",
                }
        else:
            print(f"⚠️  No manager template found for override notifications")
            results["manager_notification"] = {
                "success": False,
                "message": "Manager notification template not found",
            }

    except Exception as e:
        print(f"❌ Manager override notification error: {str(e)}")
        results["manager_notification"] = {"success": False, "message": str(e)}

    return results


# ===================================
# EMAIL TESTING FUNCTION
# ===================================


def test_email_configuration():
    """Test email configuration and connectivity"""
    try:
        if not email_config.is_configured:
            return {
                "success": False,
                "message": "Email not configured - missing SMTP credentials",
                "details": {
                    "server": email_config.smtp_server,
                    "port": email_config.smtp_port,
                    "user_set": bool(email_config.smtp_user),
                    "password_set": bool(email_config.smtp_password),
                },
            }

        # Test connection
        server = email_config.get_connection()
        server.quit()

        return {
            "success": True,
            "message": "Email configuration is working correctly",
            "details": {
                "server": email_config.smtp_server,
                "port": email_config.smtp_port,
                "from_email": email_config.from_email,
                "manager_emails": email_config.manager_emails,
            },
        }

    except Exception as e:
        return {
            "success": False,
            "message": f"Email configuration test failed: {str(e)}",
            "details": {
                "server": email_config.smtp_server,
                "port": email_config.smtp_port,
                "error": str(e),
            },
        }


# Add these functions to your app.py file to standardize date formatting


@app.template_filter("format_date")
def format_date_filter(date_value):
    """Format date as 'August 28, 2025'"""
    return format_date_for_display(date_value)


@app.template_filter("format_datetime")
def format_datetime_filter(dt_value):
    """Format datetime as 'August 28, 2025 at 3:30 PM EST'"""
    return format_datetime_for_display(dt_value)


@app.template_filter("format_time")
def format_time_filter(time_value):
    """Format time as '3:30 PM'"""
    if time_value is None:
        return "Unknown"

    if isinstance(time_value, str):
        try:
            # Handle both HH:MM:SS and HH:MM formats
            if len(time_value.split(":")) == 3:
                time_value = datetime.strptime(time_value, "%H:%M:%S").time()
            else:
                time_value = datetime.strptime(time_value, "%H:%M").time()
        except (ValueError, TypeError):
            return str(time_value)

    if isinstance(time_value, time_type):
        return time_value.strftime("%I:%M %p")

    return str(time_value)


# ===================================
# APPLICATION INITIALIZATION
# ===================================


if __name__ == "__main__":
    import sys

    # Check command line arguments for database operations
    if len(sys.argv) > 1:
        if sys.argv[1] == "reset-db":
            print("Resetting database...")
            if reset_database():
                print("Database reset successfully!")
            else:
                print("Database reset failed!")
            sys.exit()
        elif sys.argv[1] == "verify-db":
            print("Verifying database...")
            if verify_database():
                print("Database verification passed!")
            else:
                print("Database verification failed!")
            sys.exit()

    # Check if database exists and is properly initialized
    if not os.path.exists(app.config["DATABASE"]):
        print("Database does not exist. Initializing...")
        if init_db():
            print("Database initialized successfully!")
            print("\nDefault Login Credentials:")
            print("Senior Manager:       / admin123")
            print("Manager: manager@riversideequestrian.ca / admin123")
            print("Client: chloechow2016@gmail.com (no password needed)")
            print("-" * 50)
        else:
            print("Database initialization failed!")
            sys.exit(1)
    else:
        # Verify existing database
        print("Verifying existing database...")
        if not verify_database():
            print("Database verification failed. Attempting to reset...")
            if reset_database():
                print("Database reset and reinitialized successfully!")
            else:
                print("Database reset failed!")
                sys.exit(1)
        else:
            print("Database verification passed!")

    print("Starting Flask application...")
    app.run(debug=False, host="0.0.0.0", port=5000)
