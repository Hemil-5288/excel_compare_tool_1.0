from django.urls import path
from . import views

urlpatterns = [
    path("compare", views.compare_excel_api),
    path("compare/json", views.compare_json_api),
    path("compare/html", views.compare_html_api),
    path("compare/sheets", views.compare_sheets_api),
    path("compare/sheet/diff", views.compare_sheet_diff_api),
    path("compare/excel", views.compare_excel_file_api),
]