#!/usr/bin/env python3
"""
Database Creation Script for Riverside Equestrian Cancellation System
Run this script to create a fresh database with sample data.

Usage:
    python create_db.py
"""

import sqlite3
import os
from werkzeug.security import generate_password_hash

DATABASE_NAME = "cancellation_system.db"


def create_database():
    """Create the complete database with all tables and sample data"""

    # Remove existing database if it exists
    if os.path.exists(DATABASE_NAME):
        print(f"Removing existing database: {DATABASE_NAME}")
        os.remove(DATABASE_NAME)

    print(f"Creating new database: {DATABASE_NAME}")
    conn = sqlite3.connect(DATABASE_NAME)
    conn.row_factory = sqlite3.Row

    try:
        # Enable foreign key constraints
        conn.execute("PRAGMA foreign_keys = ON")
        print("✓ Foreign key constraints enabled")

        # Create all tables
        print("Creating tables...")

        # Students table
        conn.execute(
            """
            CREATE TABLE students (
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

        # Admin users table
        conn.execute(
            """
            CREATE TABLE admin_users (
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

        # Membership tiers table
        conn.execute(
            """
            CREATE TABLE membership_tiers (
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

        # Cancellations table
        conn.execute(
            """
            CREATE TABLE cancellations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id INTEGER NOT NULL,
                lesson_date DATE NOT NULL,
                lesson_time TIME NOT NULL,
                sequential_lessons TEXT,
                reschedule_requested BOOLEAN DEFAULT 0,
                reschedule_preferences TEXT,
                error_report TEXT,
                charged BOOLEAN DEFAULT 0,
                excluded BOOLEAN DEFAULT 0,
                approved_by TEXT,
                exclusion_reason TEXT,
                manager_notes TEXT,
                status TEXT DEFAULT 'pending',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (student_id) REFERENCES students (id) ON DELETE CASCADE
            )
        """
        )
        print("✓ Cancellations table created")

        # System settings table
        conn.execute(
            """
            CREATE TABLE system_settings (
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

        # Email templates table
        conn.execute(
            """
            CREATE TABLE email_templates (
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

        # System logs table
        conn.execute(
            """
            CREATE TABLE system_logs (
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

        # Create indexes
        print("Creating indexes...")
        indexes = [
            "CREATE INDEX idx_students_email ON students(email)",
            "CREATE INDEX idx_students_membership ON students(membership_level)",
            "CREATE INDEX idx_cancellations_student ON cancellations(student_id)",
            "CREATE INDEX idx_cancellations_date ON cancellations(lesson_date)",
            "CREATE INDEX idx_cancellations_created ON cancellations(created_at)",
            "CREATE INDEX idx_system_logs_user ON system_logs(user_id)",
            "CREATE INDEX idx_system_logs_action ON system_logs(action)",
            "CREATE INDEX idx_system_logs_created ON system_logs(created_at)",
        ]

        for index in indexes:
            conn.execute(index)
        print("✓ Database indexes created")

        # Insert membership tiers
        print("Inserting membership tiers...")
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
                INSERT INTO membership_tiers 
                (level, free_notices, deadline_hours, deadline_display, active, sort_order, color, description, welcome_message, policy_message, allow_sequential, allow_rescheduling, require_approval)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                tier,
            )
        print("✓ Membership tiers inserted")

        # Insert system settings
        print("Inserting system settings...")
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
                INSERT INTO system_settings (key, value, description, category, data_type)
                VALUES (?, ?, ?, ?, ?)
            """,
                setting,
            )
        print("✓ System settings inserted")

        # Insert email templates
        print("Inserting email templates...")
        email_templates = [
            (
                "client_confirmation",
                "Client Confirmation",
                "Cancellation Confirmation - {{client_name}}",
                "<p>Dear {{client_name}},</p><p>This email confirms that we have received your lesson cancellation request.</p><p><strong>Cancellation Details:</strong></p><ul><li>Lesson Date: {{lesson_date}}</li><li>Lesson Time: {{lesson_time}}</li><li>Membership Level: {{membership_tier}}</li><li>Cancellation Status: {{cancellation_status}}</li></ul><p>{{status_message}}</p><p>Current cancellation usage for this month:</p><ul><li>Free cancellations used: {{used_cancellations}} of {{allowed_cancellations}}</li><li>Remaining free cancellations: {{remaining_cancellations}}</li></ul><p>If you have any questions about this cancellation or our policy, please contact us.</p><p>Best regards,<br>The Riverside Equestrian Team</p>",
                "client",
                1,
                1,
                "normal",
                0,
                0,
                "client_name,lesson_date,lesson_time,membership_tier,cancellation_status,status_message,used_cancellations,allowed_cancellations,remaining_cancellations",
            ),
            (
                "manager_notification",
                "Manager Notification",
                "New Cancellation - {{client_name}} - {{lesson_date}}",
                "<p>A new cancellation has been submitted:</p><p><strong>Client:</strong> {{client_name}} ({{membership_tier}})</p><p><strong>Parent:</strong> {{parent_name}}</p><p><strong>Email:</strong> {{client_email}}</p><p><strong>Phone:</strong> {{client_phone}}</p><p><strong>Lesson:</strong> {{lesson_date}} at {{lesson_time}}</p><p><strong>Status:</strong> {{cancellation_status}}</p><p><strong>Submission time:</strong> {{submission_time}}</p>",
                "manager",
                1,
                1,
                "normal",
                0,
                0,
                "client_name,membership_tier,parent_name,client_email,client_phone,lesson_date,lesson_time,cancellation_status,submission_time",
            ),
            (
                "cancellation_charged",
                "Cancellation Charged",
                "Cancellation Notice - Charge Applied - {{client_name}}",
                "<p>Dear {{client_name}},</p><p>Your cancellation has been processed, and a charge has been applied to your account.</p><p><strong>Lesson Date:</strong> {{lesson_date}}</p><p><strong>Lesson Time:</strong> {{lesson_time}}</p><p><strong>Reason for charge:</strong> {{charge_reason}}</p><p>For questions about this charge, please contact us.</p><p>Best regards,<br>The Riverside Equestrian Team</p>",
                "client",
                1,
                1,
                "normal",
                0,
                0,
                "client_name,lesson_date,lesson_time,charge_reason",
            ),
        ]

        for template in email_templates:
            conn.execute(
                """
                INSERT INTO email_templates 
                (id, name, subject, body, type, active, auto_send, priority, delay_minutes, include_attachment, variables_used)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                template,
            )
        print("✓ Email templates inserted")

        # Insert admin users (password: admin123)
        print("Inserting admin users...")
        password_hash = generate_password_hash("admin123")
        admin_users = [
            (
                "admin@riversideequestrian.ca",
                password_hash,
                "senior_manager",
                "Senior",
                "Admin",
                1,
            ),
            (
                "manager@riversideequestrian.ca",
                password_hash,
                "manager",
                "Manager",
                "User",
                1,
            ),
        ]

        for user in admin_users:
            conn.execute(
                """
                INSERT INTO admin_users (email, password_hash, role, first_name, last_name, active)
                VALUES (?, ?, ?, ?, ?, ?)
            """,
                user,
            )
        print("✓ Admin users inserted")

        # Insert sample students
        print("Inserting sample students...")
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
                INSERT INTO students (first_name, last_name, parent_first, parent_last, email, phone, membership_level)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
                student,
            )
        print("✓ Sample students inserted")

        # Insert sample cancellations
        print("Inserting sample cancellations...")
        sample_cancellations = [
            (1, "2024-08-20", "14:00:00", 0, "approved", "2024-08-15 10:30:00"),
            (2, "2024-08-22", "16:00:00", 1, "charged", "2024-08-16 09:15:00"),
            (3, "2024-08-25", "15:30:00", 0, "approved", "2024-08-17 11:45:00"),
            (1, "2024-08-28", "14:00:00", 0, "approved", "2024-08-18 08:20:00"),
            (4, "2024-08-30", "13:00:00", 1, "charged", "2024-08-19 14:10:00"),
            (2, "2024-09-02", "16:00:00", 0, "pending", "2024-08-20 15:30:00"),
            (5, "2024-09-05", "11:00:00", 0, "approved", "2024-08-21 12:00:00"),
            (6, "2024-09-08", "17:00:00", 1, "charged", "2024-08-22 18:45:00"),
            (7, "2024-09-10", "10:00:00", 0, "pending", "2024-08-23 09:30:00"),
            (8, "2024-09-12", "18:00:00", 0, "approved", "2024-08-24 16:20:00"),
            (3, "2024-09-15", "15:30:00", 1, "charged", "2024-08-25 12:15:00"),
            (9, "2024-09-18", "09:00:00", 0, "approved", "2024-08-26 14:45:00"),
            (10, "2024-09-20", "13:30:00", 0, "pending", "2024-08-27 11:30:00"),
        ]

        for cancellation in sample_cancellations:
            conn.execute(
                """
                INSERT INTO cancellations (student_id, lesson_date, lesson_time, charged, status, created_at)
                VALUES (?, ?, ?, ?, ?, ?)
            """,
                cancellation,
            )
        print("✓ Sample cancellations inserted")

        # Commit all changes
        conn.commit()
        print("\n" + "=" * 50)
        print("DATABASE CREATED SUCCESSFULLY!")
        print("=" * 50)

        # Show summary
        cursor = conn.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = [row[0] for row in cursor.fetchall()]
        print(f"Created {len(tables)} tables: {', '.join(tables)}")

        cursor = conn.execute("SELECT COUNT(*) FROM students")
        student_count = cursor.fetchone()[0]
        print(f"Inserted {student_count} students")

        cursor = conn.execute("SELECT COUNT(*) FROM cancellations")
        cancellation_count = cursor.fetchone()[0]
        print(f"Inserted {cancellation_count} cancellations")

        cursor = conn.execute("SELECT COUNT(*) FROM membership_tiers")
        tier_count = cursor.fetchone()[0]
        print(f"Inserted {tier_count} membership tiers")

        cursor = conn.execute("SELECT COUNT(*) FROM admin_users")
        admin_count = cursor.fetchone()[0]
        print(f"Inserted {admin_count} admin users")

        print("\nLogin Credentials:")
        print("Senior Manager: admin@riversideequestrian.ca / admin123")
        print("Manager: manager@riversideequestrian.ca / admin123")
        print("Client: chloechow2016@gmail.com (no password needed)")
        print("=" * 50)

        return True

    except Exception as e:
        print(f"Error creating database: {str(e)}")
        conn.rollback()
        return False
    finally:
        conn.close()


if __name__ == "__main__":
    print("Riverside Equestrian - Database Creation Script")
    print("=" * 50)

    if create_database():
        print("Database creation completed successfully!")
    else:
        print("Database creation failed!")
        exit(1)
