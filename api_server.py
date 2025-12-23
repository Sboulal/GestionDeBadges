"""
Badge Management API Server
Flask REST API for integration with other applications
"""

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import sqlite3
from datetime import datetime
import json
import qrcode
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import os

app = Flask(__name__)
CORS(app)  # Enable CORS for cross-origin requests

DB_NAME = 'badges.db'

# ==================== Database Helper Functions ====================

def get_db_connection():
    """Create database connection"""
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Initialize database"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nom TEXT NOT NULL,
            prenom TEXT NOT NULL,
            valide INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS print_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            printed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id)
        )
    ''')
    
    conn.commit()
    conn.close()

# ==================== QR Code Generation ====================

def generate_qr_code(data):
    """Generate QR code image"""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    return img

def create_badge_pdf(user_data):
    """Create badge PDF with QR code"""
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    
    # Title
    c.setFont("Helvetica-Bold", 24)
    c.drawCentredString(width/2, height - 100, "BADGE D'ACCÈS")
    
    # User information
    c.setFont("Helvetica", 18)
    full_name = f"{user_data.get('first_name', user_data.get('prenom', ''))} {user_data.get('last_name', user_data.get('nom', ''))}"
    c.drawCentredString(width/2, height - 150, full_name.upper())
    
    # ID
    c.setFont("Helvetica", 12)
    c.drawCentredString(width/2, height - 180, f"ID: {user_data.get('id', 'N/A')}")
    
    # Generate QR code
    qr_data = json.dumps({
        'id': user_data.get('id', 0),
        'nom': user_data.get('last_name', user_data.get('nom', '')),
        'prenom': user_data.get('first_name', user_data.get('prenom', '')),
        'timestamp': datetime.now().isoformat()
    })
    
    qr_img = generate_qr_code(qr_data)
    
    # Save QR code to buffer
    qr_buffer = BytesIO()
    qr_img.save(qr_buffer, format='PNG')
    qr_buffer.seek(0)
    
    # Add QR code to PDF
    qr_reader = ImageReader(qr_buffer)
    qr_size = 200
    c.drawImage(qr_reader, 
               (width - qr_size) / 2, 
               height - 450, 
               width=qr_size, 
               height=qr_size)
    
    # Issue date
    c.setFont("Helvetica", 10)
    c.drawCentredString(width/2, 100, 
                      f"Émis le: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    
    c.save()
    buffer.seek(0)
    return buffer

# ==================== API Routes ====================

@app.route('/')
def index():
    """API information endpoint"""
    return jsonify({
        'name': 'Badge Management API',
        'version': '1.0',
        'endpoints': {
            'GET /api/getbadges': 'Get all badges',
            'GET /api/getbadges/<id>': 'Get badge by ID',
            'POST /api/badges': 'Create new badge',
            'PUT /api/badges/<id>': 'Update badge',
            'DELETE /api/badges/<id>': 'Delete badge',
            'POST /print-label': 'Print badge (returns PDF)',
            'POST /user_data': 'Add user data',
            'GET /api/stats': 'Get statistics',
            'POST /api/validate/<id>': 'Validate badge',
            'GET /api/search': 'Search badges'
        }
    })

@app.route('/api/getbadges', methods=['GET'])
def get_all_badges():
    """Get all badges (compatible with Angular app)"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Get filter parameters
        valide = request.args.get('valide', type=int)
        search = request.args.get('search', '')
        
        query = 'SELECT * FROM users'
        params = []
        
        conditions = []
        if valide is not None:
            conditions.append('valide = ?')
            params.append(valide)
        
        if search:
            conditions.append('(nom LIKE ? OR prenom LIKE ? OR CAST(id AS TEXT) LIKE ?)')
            params.extend([f'%{search}%', f'%{search}%', f'%{search}%'])
        
        if conditions:
            query += ' WHERE ' + ' AND '.join(conditions)
        
        query += ' ORDER BY id ASC'
        
        cursor.execute(query, params)
        users = cursor.fetchall()
        conn.close()
        
        result = []
        for user in users:
            result.append({
                'id': user['id'],
                'nom': user['nom'],
                'prenom': user['prenom'],
                'valide': user['valide'],
                'created_at': user['created_at'],
                'updated_at': user['updated_at']
            })
        
        return jsonify(result)
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/getbadges/<int:badge_id>', methods=['GET'])
def get_badge_by_id(badge_id):
    """Get specific badge by ID"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM users WHERE id = ?', (badge_id,))
        user = cursor.fetchone()
        conn.close()
        
        if user:
            return jsonify({
                'id': user['id'],
                'nom': user['nom'],
                'prenom': user['prenom'],
                'valide': user['valide'],
                'created_at': user['created_at'],
                'updated_at': user['updated_at']
            })
        else:
            return jsonify({'error': 'Badge not found'}), 404
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/badges', methods=['POST'])
def create_badge():
    """Create new badge"""
    try:
        data = request.get_json()
        
        nom = data.get('nom', data.get('last_name', ''))
        prenom = data.get('prenom', data.get('first_name', ''))
        valide = data.get('valide', 0)
        
        if not nom or not prenom:
            return jsonify({'error': 'nom and prenom are required'}), 400
        
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO users (nom, prenom, valide, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?)
        ''', (nom, prenom, valide, datetime.now(), datetime.now()))
        
        user_id = cursor.lastrowid
        conn.commit()
        conn.close()
        
        return jsonify({
            'message': 'Badge created successfully',
            'id': user_id,
            'nom': nom,
            'prenom': prenom
        }), 201
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/badges/<int:badge_id>', methods=['PUT'])
def update_badge(badge_id):
    """Update existing badge"""
    try:
        data = request.get_json()
        
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Check if badge exists
        cursor.execute('SELECT * FROM users WHERE id = ?', (badge_id,))
        if not cursor.fetchone():
            conn.close()
            return jsonify({'error': 'Badge not found'}), 404
        
        # Update fields
        nom = data.get('nom', data.get('last_name'))
        prenom = data.get('prenom', data.get('first_name'))
        valide = data.get('valide')
        
        updates = []
        params = []
        
        if nom is not None:
            updates.append('nom = ?')
            params.append(nom)
        
        if prenom is not None:
            updates.append('prenom = ?')
            params.append(prenom)
        
        if valide is not None:
            updates.append('valide = ?')
            params.append(valide)
        
        updates.append('updated_at = ?')
        params.append(datetime.now())
        
        params.append(badge_id)
        
        query = f"UPDATE users SET {', '.join(updates)} WHERE id = ?"
        cursor.execute(query, params)
        conn.commit()
        conn.close()
        
        return jsonify({'message': 'Badge updated successfully'})
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/badges/<int:badge_id>', methods=['DELETE'])
def delete_badge(badge_id):
    """Delete badge"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT * FROM users WHERE id = ?', (badge_id,))
        if not cursor.fetchone():
            conn.close()
            return jsonify({'error': 'Badge not found'}), 404
        
        cursor.execute('DELETE FROM users WHERE id = ?', (badge_id,))
        cursor.execute('DELETE FROM print_logs WHERE user_id = ?', (badge_id,))
        conn.commit()
        conn.close()
        
        return jsonify({'message': 'Badge deleted successfully'})
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/print-label', methods=['POST'])
def print_label():
    """Print badge label (compatible with Angular app)"""
    try:
        data = request.get_json()
        
        # Support both formats
        nom = data.get('nom', data.get('last_name', ''))
        prenom = data.get('prenom', data.get('first_name', ''))
        user_id = data.get('id')
        
        if not nom or not prenom:
            return jsonify({'error': 'last_name and first_name are required'}), 400
        
        # If no ID provided, create new user
        if not user_id:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO users (nom, prenom, valide, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?)
            ''', (nom, prenom, 1, datetime.now(), datetime.now()))
            user_id = cursor.lastrowid
            conn.commit()
            conn.close()
        
        # Create user data for PDF
        user_data = {
            'id': user_id,
            'nom': nom,
            'prenom': prenom,
            'last_name': nom,
            'first_name': prenom
        }
        
        # Generate PDF
        pdf_buffer = create_badge_pdf(user_data)
        
        # Log print
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO print_logs (user_id, printed_at)
            VALUES (?, ?)
        ''', (user_id, datetime.now()))
        conn.commit()
        conn.close()
        
        return send_file(
            pdf_buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f'badge_{prenom}_{nom}.pdf'
        )
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/user_data', methods=['POST'])
def user_data():
    """Add user data endpoint (compatible with Angular app)"""
    try:
        data = request.get_json()
        
        nom = data.get('last_name', '')
        prenom = data.get('first_name', '')
        
        if not nom or not prenom:
            return jsonify({'error': 'last_name and first_name are required'}), 400
        
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO users (nom, prenom, valide, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?)
        ''', (nom, prenom, 1, datetime.now(), datetime.now()))
        
        user_id = cursor.lastrowid
        conn.commit()
        conn.close()
        
        return jsonify({
            'message': 'User data added successfully',
            'id': user_id
        }), 201
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/validate/<int:badge_id>', methods=['POST'])
def validate_badge(badge_id):
    """Validate/invalidate badge"""
    try:
        data = request.get_json()
        valide = data.get('valide', 1)
        
        conn = get_db_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT * FROM users WHERE id = ?', (badge_id,))
        if not cursor.fetchone():
            conn.close()
            return jsonify({'error': 'Badge not found'}), 404
        
        cursor.execute('''
            UPDATE users SET valide = ?, updated_at = ?
            WHERE id = ?
        ''', (valide, datetime.now(), badge_id))
        conn.commit()
        conn.close()
        
        return jsonify({'message': 'Badge validation status updated'})
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/search', methods=['GET'])
def search_badges():
    """Search badges"""
    try:
        query = request.args.get('q', '')
        
        if not query:
            return jsonify({'error': 'Query parameter q is required'}), 400
        
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT * FROM users 
            WHERE nom LIKE ? OR prenom LIKE ? OR CAST(id AS TEXT) LIKE ?
            ORDER BY id ASC
        ''', (f'%{query}%', f'%{query}%', f'%{query}%'))
        
        users = cursor.fetchall()
        conn.close()
        
        result = []
        for user in users:
            result.append({
                'id': user['id'],
                'nom': user['nom'],
                'prenom': user['prenom'],
                'valide': user['valide'],
                'created_at': user['created_at'],
                'updated_at': user['updated_at']
            })
        
        return jsonify(result)
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/stats', methods=['GET'])
def get_statistics():
    """Get statistics"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Total users
        cursor.execute('SELECT COUNT(*) as total FROM users')
        total = cursor.fetchone()['total']
        
        # Validated users
        cursor.execute('SELECT COUNT(*) as validated FROM users WHERE valide = 1')
        validated = cursor.fetchone()['validated']
        
        # Total prints
        cursor.execute('SELECT COUNT(*) as prints FROM print_logs')
        prints = cursor.fetchone()['prints']
        
        # Recent users (last 24h)
        cursor.execute('''
            SELECT COUNT(*) as recent 
            FROM users 
            WHERE created_at >= datetime('now', '-1 day')
        ''')
        recent = cursor.fetchone()['recent']
        
        conn.close()
        
        return jsonify({
            'total_badges': total,
            'validated_badges': validated,
            'non_validated_badges': total - validated,
            'total_prints': prints,
            'recent_badges_24h': recent
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/bulk-import', methods=['POST'])
def bulk_import():
    """Bulk import badges from JSON"""
    try:
        data = request.get_json()
        users = data.get('users', [])
        
        if not users or not isinstance(users, list):
            return jsonify({'error': 'users array is required'}), 400
        
        conn = get_db_connection()
        cursor = conn.cursor()
        
        imported = 0
        errors = []
        
        for user in users:
            nom = user.get('nom', user.get('Nom', user.get('last_name', '')))
            prenom = user.get('prenom', user.get('Prénom', user.get('first_name', '')))
            
            if nom and prenom:
                try:
                    cursor.execute('''
                        INSERT INTO users (nom, prenom, valide, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?)
                    ''', (nom, prenom, 0, datetime.now(), datetime.now()))
                    imported += 1
                except Exception as e:
                    errors.append(f"Error importing {prenom} {nom}: {str(e)}")
        
        conn.commit()
        conn.close()
        
        return jsonify({
            'message': f'{imported} badges imported successfully',
            'imported': imported,
            'errors': errors
        }), 201
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ==================== Error Handlers ====================

@app.errorhandler(404)
def not_found(error):
    return jsonify({'error': 'Endpoint not found'}), 404

@app.errorhandler(500)
def internal_error(error):
    return jsonify({'error': 'Internal server error'}), 500

# ==================== Run Server ====================

if __name__ == '__main__':
    init_db()
    print("=" * 60)
    print("Badge Management API Server")
    print("=" * 60)
    print("Server starting on http://127.0.0.1:5000")
    print("\nAvailable endpoints:")
    print("  GET  /api/getbadges       - Get all badges")
    print("  GET  /api/getbadges/<id>  - Get badge by ID")
    print("  POST /api/badges          - Create new badge")
    print("  PUT  /api/badges/<id>     - Update badge")
    print("  DELETE /api/badges/<id>   - Delete badge")
    print("  POST /print-label         - Print badge")
    print("  POST /user_data           - Add user data")
    print("  GET  /api/stats           - Get statistics")
    print("  POST /api/validate/<id>   - Validate badge")
    print("  GET  /api/search?q=...    - Search badges")
    print("  POST /api/bulk-import     - Bulk import")
    print("=" * 60)
    
    app.run(host='0.0.0.0', port=5000, debug=True)