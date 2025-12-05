import os
from datetime import timedelta

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "S4C-Processed-Documents")
COMMON_MACRO_FOLDER = os.path.join(BASE_DIR, "S4c-Macros")
DEFAULT_MACRO_NAME = 'CE_Tool.dotm'
REPORT_FOLDER = "reports"
DATABASE = os.path.join(BASE_DIR, "reference_validator.db")
LOG_FILE = os.path.join(BASE_DIR, 'user_activity.log')

ALLOWED_EXTENSIONS = {'.doc', '.docx'}
WORD_START_RETRIES = 3
TOKEN_TTL = timedelta(hours=1)

ROUTE_MACROS = {
    'language': {
        'name': 'Language Editing',
        'description': 'Language editing and grammar correction tools',
        'icon': 'edit',
        'macros': [
            "LanguageEdit.GrammarCheck_WithErrorHandling",
            "LanguageEdit.SpellCheck_Advanced",
            "LanguageEdit.StyleConsistency_Check",
            "LanguageEdit.ReadabilityAnalysis",
            "LanguageEdit.TerminologyValidation"
        ]
    },
    'technical': {
        'name': 'Technical Editing',
        'description': 'Technical document formatting and validation tools',
        'icon': 'cog',
        'macros': [
            "Referencevalidation.ValidateBWNumCite_WithErrorHandling",
            "ReferenceRenumber.Reorderbasedonseq",
            "Copyduplicate.duplicate4",
            "citationupdateonly.citationupdate",
            "techinal.technicalhighlight"
        ]
    },
    'macro_processing': {
        'name': 'Reference Processing',
        'description': 'Reference validation and citation tools',
        'icon': 'bookmark',
        'macros': [
            "Referencevalidation.ValidateBWNumCite_WithErrorHandling",
            "ReferenceRenumber.Reorderbasedonseq",
            "Copyduplicate.duplicate4",
            "citationupdateonly.citationupdate",
            "Prediting.Preditinghighlight",
            "msrpre.GenerateDashboardReport",
        ]
    },
    'ppd': {
        'name': 'PPD Processing',
        'description': 'PPD final processing tools (from PPD_Final.py)',
        'icon': 'magic',
        'macros': [
            "PPD_HTML.GenerateDocument",
            "PPD_HTML.Generate_HTML_WORDReport"
        ]
    }
}

ROUTE_PERMISSIONS = {
    'language': ['COPYEDIT', 'PM', 'ADMIN'],
    'technical': ['COPYEDIT', 'PM', 'ADMIN'],
    'macro_processing': ['COPYEDIT', 'PPD', 'PM', 'ADMIN'],
    'ppd': ['PPD', 'PM', 'ADMIN']
}
