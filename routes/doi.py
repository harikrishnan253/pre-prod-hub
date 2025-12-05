from flask import Blueprint, render_template, session, flash, redirect, url_for

doi_bp = Blueprint('doi', __name__)

@doi_bp.route('/doi_finder')
def doi_finder():
    """DOI Correction and Metadata Finder"""
    if 'user_id' not in session:
        flash("Please log in to continue.")
        return redirect(url_for('auth.login'))

    return render_template('doi_finder.html')
