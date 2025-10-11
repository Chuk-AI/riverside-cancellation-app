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

# Email System Implementation for Riverside Equestrian
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import threading


load_dotenv()
# ===================================
# APPLICATION SETUP
# ===================================

app = Flask(__name__)
app.secret_key = "riverside-equestrian-secret-key-change-in-production"
app.config["DATABASE"] = "cancellation_system.db"


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
            ("timezone", "America/Vancouver", "System timezone", "general", "string"),
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
    if not month:
        month = datetime.now().month
    if not year:
        year = datetime.now().year

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


def will_be_charged(student, lesson_datetime):
    """Check if a cancellation will be charged"""
    tier = get_membership_tier(student["membership_level"])
    status = calculate_cancellation_status(student)

    # Check if within deadline
    now = datetime.now()
    time_diff = lesson_datetime - now
    hours_diff = time_diff.total_seconds() / 3600

    if hours_diff < tier["deadline_hours"]:
        return True, "Notice will be received after the cancellation cutoff"

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
            return redirect(url_for("dashboard"))

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
            return redirect(url_for("dashboard"))

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
            cancellation_dict["lesson_time"] = datetime.strptime(
                cancellation["lesson_time"], "%H:%M:%S"
            ).time()
            cancellation_dict["created_at"] = datetime.strptime(
                cancellation["created_at"], "%Y-%m-%d %H:%M:%S"
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

        if lesson_datetime <= datetime.now():
            flash("Cannot cancel lessons that have already occurred", "error")
            return redirect(url_for("client_cancel"))

        # Check if will be charged
        will_charge, charge_reason = will_be_charged(client, lesson_datetime)

        # Determine deadline status for new database fields
        tier = get_membership_tier(client["membership_level"])
        deadline_hours = tier["deadline_hours"] if tier else 18
        hours_until_lesson = (lesson_datetime - datetime.now()).total_seconds() / 3600
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
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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
    current_time = datetime.now()

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
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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
    """Process individual cancellation - UPDATED with override tracking"""
    data = request.json
    action = data.get("action")  # 'approve' or 'charge'
    cancellation_id = data.get("cancellation_id")
    reason = data.get("reason", "")

    if not action or not cancellation_id:
        return jsonify({"success": False, "message": "Missing required fields"})

    if not reason.strip():
        return jsonify({"success": False, "message": "Override reason is required"})

    conn = get_db()

    try:
        if action == "approve":
            # Mark as approved (free cancellation) with override flag
            conn.execute(
                """UPDATE cancellations 
                   SET status = 'approved', charged = 0, manager_notes = ?, 
                       is_override = 1, updated_at = ? 
                   WHERE id = ?""",
                (
                    f"Manager Override: {reason}",
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    cancellation_id,
                ),
            )
            log_message = (
                f"Cancellation {cancellation_id} approved as free (Override: {reason})"
            )

        elif action == "charge":
            # Mark as charged with override flag
            conn.execute(
                """UPDATE cancellations 
                   SET charged = 1, status = 'charged', manager_notes = ?, 
                       is_override = 1, updated_at = ? 
                   WHERE id = ?""",
                (
                    f"Manager Override: {reason}",
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    cancellation_id,
                ),
            )
            log_message = (
                f"Cancellation {cancellation_id} marked as charged (Override: {reason})"
            )

        else:
            return jsonify({"success": False, "message": "Invalid action"})

        conn.commit()
        conn.close()

        log_action("cancellation_processed", log_message)
        return jsonify({"success": True})

    except Exception as e:
        conn.close()
        return jsonify({"success": False, "message": str(e)})


@app.route("/manager/api/cancellation/batch", methods=["POST"])
@login_required
@admin_required
def batch_process_cancellations():
    """Batch process multiple cancellations - UPDATED with override tracking"""
    data = request.json
    action = data.get("action")  # 'approve', 'charge', or 'exclude'
    cancellation_ids = data.get("cancellation_ids", [])
    reason = data.get("reason", "Batch processing")

    if not action or not cancellation_ids:
        return jsonify({"success": False, "message": "Missing required fields"})

    conn = get_db()

    try:
        processed_count = 0

        for cancellation_id in cancellation_ids:
            if action == "approve":
                conn.execute(
                    """UPDATE cancellations 
                       SET status = 'approved', charged = 0, manager_notes = ?, 
                           is_override = 1, updated_at = ? 
                       WHERE id = ?""",
                    (
                        f"Batch approval: {reason}",
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        cancellation_id,
                    ),
                )
            elif action == "charge":
                conn.execute(
                    """UPDATE cancellations 
                       SET charged = 1, status = 'charged', manager_notes = ?, 
                           is_override = 1, updated_at = ? 
                       WHERE id = ?""",
                    (
                        f"Batch charge: {reason}",
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        cancellation_id,
                    ),
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
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        cancellation_id,
                    ),
                )

            processed_count += 1

        conn.commit()
        conn.close()

        log_action(
            "batch_processing",
            f"Batch {action}: {processed_count} cancellations ({reason})",
        )
        return jsonify({"success": True, "processed": processed_count})

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
                f"Reverted by {session['user_email']} on {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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
        new_notes = f"{existing_notes}\n[{datetime.now().strftime('%Y-%m-%d %H:%M')} - {session['user_email']}]: {note}".strip()

        conn.execute(
            "UPDATE cancellations SET manager_notes = ?, updated_at = ? WHERE id = ?",
            (new_notes, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), cancellation_id),
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
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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

        # Get cancellation statistics
        stats = conn.execute(
            """
            SELECT 
                COUNT(*) as total_cancellations,
                SUM(CASE WHEN strftime('%Y-%m', created_at) = strftime('%Y-%m', 'now') THEN 1 ELSE 0 END) as this_month,
                SUM(CASE WHEN charged = 0 AND excluded = 0 AND strftime('%Y-%m', created_at) = strftime('%Y-%m', 'now') THEN 1 ELSE 0 END) as free_used
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


def get_analytics_summary(date_range="month", status_filter="", membership_filter=""):
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

        date_clause = date_conditions.get(date_range, date_conditions["month"])
        where_clause = f"WHERE {date_clause}"
        params = []

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
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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


@app.route("/manager/api/student/<int:student_id>/details", methods=["GET"])
@login_required
@admin_required
def get_student_details_for_edit(student_id):
    """Get student details for editing (different from the existing one)"""
    try:
        conn = get_db()

        student = conn.execute(
            "SELECT * FROM students WHERE id = ?", (student_id,)
        ).fetchone()

        if not student:
            return jsonify({"success": False, "message": "Student not found"})

        conn.close()

        return jsonify({"success": True, "student": dict(student)})

    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


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
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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
            (new_membership, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), student_id),
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


@app.route("/senior/templates")
@login_required
@senior_admin_required
def senior_templates():
    """Senior manager email templates"""
    conn = get_db()

    # Get all email templates
    templates = conn.execute("SELECT * FROM email_templates ORDER BY name").fetchall()

    conn.close()

    return render_template("senior_templates.html", templates=templates)


@app.route("/senior/delete-template/<template_id>", methods=["DELETE"])
@login_required
@senior_admin_required
def delete_template(template_id):
    """Delete email template"""
    try:
        conn = get_db()
        conn.execute("DELETE FROM email_templates WHERE id = ?", (template_id,))
        conn.commit()
        conn.close()

        log_action("template_deleted", f"Deleted template: {template_id}")
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/senior/import-templates", methods=["POST"])
@login_required
@senior_admin_required
def import_templates():
    """Import email templates"""
    try:
        data = request.json
        templates = data.get("templates", [])

        conn = get_db()
        imported_count = 0

        for template in templates:
            conn.execute(
                """
                INSERT OR REPLACE INTO email_templates 
                (id, name, subject, body, type, active, auto_send, priority, delay_minutes, include_attachment)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    template.get("id"),
                    template.get(
                        "title", template.get("id", "").replace("_", " ").title()
                    ),
                    template.get("subject", ""),
                    template.get("body", ""),
                    template.get("type", "client"),
                    template.get("active", True),
                    template.get("autoSend", True),
                    template.get("priority", "normal"),
                    template.get("delay", 0),
                    template.get("includeAttachment", False),
                ),
            )
            imported_count += 1

        conn.commit()
        conn.close()

        log_action("templates_imported", f"Imported {imported_count} templates")
        return jsonify({"success": True, "count": imported_count})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


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


@app.route("/senior/send-test-email", methods=["POST"])
@login_required
@senior_admin_required
def send_test_email_route():
    """Send test email"""
    data = request.json
    template_id = data.get("template_id")
    test_email = data.get("test_email")

    if not template_id or not test_email:
        return jsonify({"success": False, "error": "Missing template_id or test_email"})

    result = send_test_email(template_id, test_email)
    return jsonify(result)


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
    return {"now": datetime.now()}


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


# ===================================
# EMAIL CONFIGURATION
# ===================================
class EmailConfig:
    """Email configuration class"""

    def __init__(self):
        # Environment variables for production
        self.smtp_server = os.getenv(
            "SMTP_SERVER", "sandbox.smtp.mailtrap.io"
        )  # Mailtrap for development
        self.smtp_port = int(os.getenv("SMTP_PORT", "2525"))
        self.smtp_user = os.getenv("SMTP_USER", "your_mailtrap_username")
        self.smtp_password = os.getenv("SMTP_PASSWORD", "your_mailtrap_password")
        self.from_email = os.getenv("FROM_EMAIL", "noreply@riversideequestrian.ca")
        self.from_name = os.getenv("FROM_NAME", "Riverside Equestrian")
        self.use_tls = os.getenv("USE_TLS", "True").lower() == "true"

        # For development/testing
        self.debug_mode = os.getenv("EMAIL_DEBUG", "True").lower() == "true"
        self.log_emails = True


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


def log_email_attempt(to_email, subject, success, error=None):
    """Log email sending attempts"""
    try:
        conn = get_db()
        conn.execute(
            """
            INSERT INTO system_logs (user_id, user_type, action, details, created_at)
            VALUES (?, ?, ?, ?, ?)
        """,
            (
                session.get("user_id"),
                "system",
                "email_sent" if success else "email_failed",
                f"To: {to_email}, Subject: {subject}"
                + (f", Error: {error}" if error else ""),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Failed to log email attempt: {e}")


def send_email(to_email, subject, body, template_type="client", attachments=None):
    """
    Send email using SMTP

    Args:
        to_email (str): Recipient email address
        subject (str): Email subject
        body (str): HTML email body
        template_type (str): Type of email template
        attachments (list): List of file paths to attach

    Returns:
        dict: Result with success status and message
    """
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
                    print(f"Failed to attach file {file_path}: {e}")

        # Send email
        server = smtplib.SMTP(email_config.smtp_server, email_config.smtp_port)

        if email_config.use_tls:
            server.starttls()

        if email_config.smtp_user and email_config.smtp_password:
            server.login(email_config.smtp_user, email_config.smtp_password)

        server.send_message(msg)
        server.quit()

        result = {"success": True, "message": "Email sent successfully"}

        if email_config.debug_mode:
            print(f"✓ Email sent to {to_email}: {subject}")

        log_email_attempt(to_email, subject, True)
        return result

    except Exception as e:
        error_msg = str(e)
        result = {"success": False, "message": f"Failed to send email: {error_msg}"}

        if email_config.debug_mode:
            print(f"✗ Email failed to {to_email}: {error_msg}")

        log_email_attempt(to_email, subject, False, error_msg)
        return result


@send_email_async
def send_email_async_wrapper(
    to_email, subject, body, template_type="client", attachments=None
):
    """Async wrapper for sending emails"""
    return send_email(to_email, subject, body, template_type, attachments)


# ===================================
# TEMPLATE PROCESSING FUNCTIONS
# ===================================


def get_email_template(template_id):
    """Get email template from database"""
    try:
        conn = get_db()
        template = conn.execute(
            "SELECT * FROM email_templates WHERE id = ? AND active = 1", (template_id,)
        ).fetchone()
        conn.close()
        return dict(template) if template else None
    except Exception as e:
        print(f"Error getting template {template_id}: {e}")
        return None


def process_template_variables(template_body, template_subject, variables):
    """
    Replace template variables with actual values

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


# Add these functions to your app.py file to standardize date formatting


@app.template_filter("format_date")
def format_date_filter(date_value):
    """Format date as 'August 28, 2025'"""
    if date_value is None:
        return "Unknown"

    if isinstance(date_value, str):
        try:
            if "T" in date_value:  # ISO format with time
                date_value = datetime.fromisoformat(date_value.replace("Z", "+00:00"))
            elif len(date_value) == 10:  # YYYY-MM-DD format
                date_value = datetime.strptime(date_value, "%Y-%m-%d")
            else:  # Try parsing as datetime string
                date_value = datetime.strptime(date_value, "%Y-%m-%d %H:%M:%S")
        except (ValueError, TypeError):
            return str(date_value)
    elif isinstance(date_value, date_type) and not isinstance(date_value, datetime):
        # Convert date to datetime for consistent formatting
        date_value = datetime.combine(date_value, time_type())

    if isinstance(date_value, datetime):
        return date_value.strftime("%B %d, %Y")

    return str(date_value)


@app.template_filter("format_datetime")
def format_datetime_filter(dt_value):
    """Format datetime as 'August 28, 2025 at 3:30 PM'"""
    if dt_value is None:
        return "Unknown"

    if isinstance(dt_value, str):
        try:
            if "T" in dt_value:
                dt_value = datetime.fromisoformat(dt_value.replace("Z", "+00:00"))
            else:
                dt_value = datetime.strptime(dt_value, "%Y-%m-%d %H:%M:%S")
        except (ValueError, TypeError):
            return str(dt_value)

    if isinstance(dt_value, datetime):
        return dt_value.strftime("%B %d, %Y at %I:%M %p")

    return str(dt_value)


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


# Update the get_template_variables function to use consistent formatting
def get_template_variables(student=None, cancellation=None, extra_vars=None):
    """Generate template variables with consistent date formatting"""
    variables = {
        "current_date": datetime.now().strftime("%B %d, %Y"),
        "current_time": datetime.now().strftime("%I:%M %p"),
        "company_name": "Riverside Equestrian",
        "contact_email": "managers@riversideequestrian.ca",
        "policy_url": "https://www.riversideequestrian.ca/cancellations",
        "website_url": "https://www.riversideequestrian.ca",
    }

    if student:
        variables.update(
            {
                "client_name": f"{student['first_name']} {student['last_name']}",
                "client_first_name": student["first_name"],
                "client_last_name": student["last_name"],
                "client_email": student["email"],
                "client_phone": student.get("phone", ""),
                "membership_tier": student["membership_level"],
                "parent_name": f"{student.get('parent_first', '')} {student.get('parent_last', '')}".strip(),
                "parent_first_name": student.get("parent_first", ""),
                "parent_last_name": student.get("parent_last", ""),
            }
        )

        # Get membership tier info
        tier = get_membership_tier(student["membership_level"])
        if tier:
            variables.update(
                {
                    "allowed_cancellations": str(tier["free_notices"]),
                    "cancellation_deadline": tier["deadline_display"],
                }
            )

        # Get current usage
        status = calculate_cancellation_status(student)
        variables.update(
            {
                "used_cancellations": str(status["used"]),
                "remaining_cancellations": str(status["remaining"]),
            }
        )

    if cancellation:
        # Format dates consistently
        if isinstance(cancellation["lesson_date"], str):
            lesson_date = datetime.strptime(cancellation["lesson_date"], "%Y-%m-%d")
        else:
            lesson_date = cancellation["lesson_date"]

        if isinstance(cancellation["lesson_time"], str):
            try:
                lesson_time = datetime.strptime(
                    cancellation["lesson_time"], "%H:%M:%S"
                ).time()
            except ValueError:
                lesson_time = datetime.strptime(
                    cancellation["lesson_time"], "%H:%M"
                ).time()
        else:
            lesson_time = cancellation["lesson_time"]

        variables.update(
            {
                "lesson_date": lesson_date.strftime("%B %d, %Y"),
                "lesson_date_short": lesson_date.strftime(
                    "%B %d, %Y"
                ),  # Also use full format
                "lesson_time": lesson_time.strftime("%I:%M %p"),
                "lesson_time_24h": lesson_time.strftime("%H:%M"),
                "cancellation_status": (
                    "Free cancellation"
                    if not cancellation.get("charged")
                    else "Charged cancellation"
                ),
                "will_be_charged": "Yes" if cancellation.get("charged") else "No",
                "charge_reason": get_charge_reason(cancellation),
                "submission_time": datetime.now().strftime("%B %d, %Y at %I:%M %p"),
            }
        )

        # Sequential lessons with consistent formatting
        if cancellation.get("sequential_lessons"):
            try:
                sequential = (
                    eval(cancellation["sequential_lessons"])
                    if isinstance(cancellation["sequential_lessons"], str)
                    else cancellation["sequential_lessons"]
                )
                if sequential:
                    formatted_lessons = []
                    for lesson in sequential:
                        lesson_date_str = lesson["date"]
                        lesson_time_str = lesson["time"]

                        # Parse and format date
                        if isinstance(lesson_date_str, str):
                            lesson_date_obj = datetime.strptime(
                                lesson_date_str, "%Y-%m-%d"
                            )
                        else:
                            lesson_date_obj = lesson_date_str

                        # Parse and format time
                        if isinstance(lesson_time_str, str):
                            try:
                                lesson_time_obj = datetime.strptime(
                                    lesson_time_str, "%H:%M:%S"
                                ).time()
                            except ValueError:
                                lesson_time_obj = datetime.strptime(
                                    lesson_time_str, "%H:%M"
                                ).time()
                        else:
                            lesson_time_obj = lesson_time_str

                        formatted_lessons.append(
                            f"{lesson_date_obj.strftime('%B %d, %Y')} at {lesson_time_obj.strftime('%I:%M %p')}"
                        )

                    variables["sequential_lessons"] = (
                        f"Additional lessons: {', '.join(formatted_lessons)}"
                    )
                else:
                    variables["sequential_lessons"] = "No additional lessons"
            except:
                variables["sequential_lessons"] = "No additional lessons"
        else:
            variables["sequential_lessons"] = "No additional lessons"

        # Reschedule info
        variables.update(
            {
                "reschedule_requested": (
                    "Yes" if cancellation.get("reschedule_requested") else "No"
                ),
                "reschedule_preferences": cancellation.get(
                    "reschedule_preferences", "None provided"
                ),
                "error_report": cancellation.get("error_report", "None reported"),
            }
        )

    # Add extra variables
    if extra_vars:
        variables.update(extra_vars)

    return variables


def get_charge_reason(cancellation):
    """Get human-readable reason for cancellation charge"""
    if not cancellation.get("charged"):
        return "Within policy - no charge applied"

    if cancellation.get("manager_notes"):
        return cancellation["manager_notes"]

    # Default reasons based on common scenarios
    return "Late cancellation or monthly limit exceeded"


# ===================================
# EMAIL SENDING WRAPPER FUNCTIONS
# ===================================


def send_cancellation_confirmation(student, cancellation):
    """Send cancellation confirmation email to client"""
    template = get_email_template("client_confirmation")
    if not template:
        print("Warning: No client_confirmation template found")
        return {"success": False, "message": "Template not found"}

    variables = get_template_variables(student, cancellation)

    # Add status message
    if cancellation.get("charged"):
        variables["status_message"] = (
            f"A charge will be applied to your account. Reason: {get_charge_reason(cancellation)}"
        )
    else:
        variables["status_message"] = (
            "This cancellation has been processed at no charge."
        )

    body, subject = process_template_variables(
        template["body"], template["subject"], variables
    )

    # Send to client
    result = send_email(student["email"], subject, body, "client")

    # Also send to managers if auto-send is enabled for manager notifications
    manager_template = get_email_template("manager_notification")
    if manager_template and manager_template.get("auto_send"):
        send_manager_notification(student, cancellation)

    return result


def send_manager_notification(student, cancellation):
    """Send new cancellation notification to managers"""
    template = get_email_template("manager_notification")
    if not template:
        return {"success": False, "message": "Template not found"}

    variables = get_template_variables(
        student,
        cancellation,
        {
            "action_required": "Review cancellation and approve/charge as needed",
            "dashboard_url": f"{request.url_root}manager/cancellations?student={student['id']}",
        },
    )

    body, subject = process_template_variables(
        template["body"], template["subject"], variables
    )

    # Get manager email from settings
    manager_email = get_system_setting(
        "company_email", "managers@riversideequestrian.ca"
    )

    return send_email(manager_email, subject, body, "manager")


def send_status_update_email(student, cancellation, status_type="charged"):
    """Send status update email when cancellation status changes"""
    template_mapping = {
        "charged": "cancellation_charged",
        "free": "free_cancellation",
        "excluded": "illness_exclusion",
        "late": "late_cancellation",
        "limit_exceeded": "limit_exceeded",
    }

    template_id = template_mapping.get(status_type, "client_confirmation")
    template = get_email_template(template_id)

    if not template:
        return {"success": False, "message": f"Template {template_id} not found"}

    variables = get_template_variables(student, cancellation)

    # Add specific variables based on status type
    if status_type == "charged":
        variables["charge_amount"] = "As per your membership agreement"
    elif status_type == "excluded":
        variables["documentation_date"] = datetime.now().strftime("%B %d, %Y")

    body, subject = process_template_variables(
        template["body"], template["subject"], variables
    )

    return send_email(student["email"], subject, body, "client")


def send_test_email(template_id, test_email_address):
    """Send test email with sample data"""
    template = get_email_template(template_id)
    if not template:
        return {"success": False, "message": "Template not found"}

    # Sample data for testing
    sample_student = {
        "first_name": "Sarah",
        "last_name": "Johnson",
        "parent_first": "Michael",
        "parent_last": "Johnson",
        "email": test_email_address,
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

    variables = get_template_variables(
        sample_student,
        sample_cancellation,
        {
            "status_message": "This is a test email with sample data.",
        },
    )

    body, subject = process_template_variables(
        template["body"], template["subject"], variables
    )

    # Add test disclaimer
    test_subject = f"[TEST] {subject}"
    test_body = f"""
    <div style="background-color: #fff3cd; border: 1px solid #ffeaa7; padding: 10px; margin-bottom: 20px; border-radius: 5px;">
        <strong>⚠️ TEST EMAIL</strong><br>
        This is a test email with sample data. This would be sent to actual clients in production.
    </div>
    {body}
    """

    return send_email(test_email_address, test_subject, test_body, "test")


def get_system_setting(key, default=None):
    """Get system setting from database"""
    try:
        conn = get_db()
        setting = conn.execute(
            "SELECT value FROM system_settings WHERE key = ?", (key,)
        ).fetchone()
        conn.close()
        return setting["value"] if setting else default
    except:
        return default


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
