import json
from django.http import JsonResponse, FileResponse
from django.views.decorators.csrf import csrf_exempt
from django.shortcuts import render
from io import BytesIO
from .excel_comparator_india import ( 
    compare_excel_with_gain_summary_inline_India,
    write_html_report_India,
    generate_comparison_excel_India
)

from .excel_comparator import (
    write_html_report,
    compare_excel_with_gain_summary_inline,
    compare_excel_sheets_inline,
    compare_single_sheet_diff_inline,
    generate_comparison_excel,
)


def home(request):
    return render(request, "index.html")


# API Endpoints
@csrf_exempt
def compare_excel_api(request):
    if request.method != "POST":
        return JsonResponse({"error": "Invalid request"}, status=400)
    try:
        original_file = request.FILES.get("original_file")
        website_file = request.FILES.get("website_file")
        sheets_config = request.POST.get("sheets_config")
        country = request.POST.get("country", "usa")

        parsed_config = None
        if sheets_config:
            try:
                parsed_config = json.loads(sheets_config)
            except Exception:
                parsed_config = None

        if country == "india":
            from .excel_comparator_india import compare_excel_with_gain_summary_inline_India
            result = compare_excel_with_gain_summary_inline_India(
                original_file.read(),
                website_file.read(),
                sheets_config=parsed_config,
            )
        else:
            result = compare_excel_with_gain_summary_inline(
                original_file.read(),
                website_file.read(),
                sheets_config=parsed_config,
            )
        # Remove excel_bytes from result before returning JSON
        if "excel_bytes" in result:
            del result["excel_bytes"]
        return JsonResponse(result)
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)


@csrf_exempt
def compare_json_api(request):
    if request.method != "POST":
        return JsonResponse({"error": "Invalid request"}, status=400)
    try:
        original_file = request.FILES.get("original_file")
        website_file = request.FILES.get("website_file")
        sheets_config = request.POST.get("sheets_config")
        country = request.POST.get("country", "usa")

        parsed_config = None
        if sheets_config:
            try:
                parsed_config = json.loads(sheets_config)
            except Exception:
                parsed_config = None

        if country == "india":
            result = compare_excel_with_gain_summary_inline_India(
                original_file.read(),
                website_file.read(),
                sheets_config=parsed_config,
            )
        else:
            result = compare_excel_with_gain_summary_inline(
                original_file.read(),
                website_file.read(),
                sheets_config=parsed_config,
            )

        return JsonResponse({"results": result["summary_rows"]})
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)


@csrf_exempt
def compare_html_api(request):
    if request.method != "POST":
        return JsonResponse({"error": "Invalid request"}, status=400)
    try:
        original_file = request.FILES.get("original_file")
        website_file = request.FILES.get("website_file")
        sheets_config = request.POST.get("sheets_config")
        country = request.POST.get("country", "usa")

        parsed_config = None
        if sheets_config:
            try:
                parsed_config = json.loads(sheets_config)
            except Exception:
                parsed_config = None

        if country == "india":
            result = compare_excel_with_gain_summary_inline_India(
                original_file.read(),
                website_file.read(),
                sheets_config=parsed_config,
            )
            html = write_html_report_India(result["summary_rows"])
            return JsonResponse({"html": html})
        else:
            result = compare_excel_with_gain_summary_inline(
                original_file.read(),
                website_file.read(),
                sheets_config=parsed_config,
            )
            html = write_html_report(result["summary_rows"])
            return JsonResponse({"html": html})
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)


@csrf_exempt
def compare_sheets_api(request):
    if request.method != "POST":
        return JsonResponse({"error": "Invalid request"}, status=400)
    try:
        original_file = request.FILES.get("original_file")
        website_file = request.FILES.get("website_file")
        country = request.POST.get("country", "usa")

        if country == "india":
            result = compare_excel_with_gain_summary_inline_India(
                original_file.read(),
                website_file.read(),
            )
            sheets_with_diff = []
            for sname, sdata in result["sheets"].items():
                if (len(sdata.get("different_rows", [])) > 1 or
                    len(sdata.get("only_in_original_rows", [])) > 1 or
                    len(sdata.get("only_in_website_rows", [])) > 1):
                    sheets_with_diff.append(sname)
            return JsonResponse({"sheets": sheets_with_diff})
        else:
            result = compare_excel_sheets_inline(
                original_file.read(),
                website_file.read(),
            )
            return JsonResponse(result)
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)


@csrf_exempt
def compare_sheet_diff_api(request):
    if request.method != "POST":
        return JsonResponse({"error": "Invalid request"}, status=400)
    try:
        original_file = request.FILES.get("original_file")
        website_file = request.FILES.get("website_file")
        sheet_name = request.POST.get("sheet_name")
        country = request.POST.get("country", "usa")

        if country == "india":
            result = compare_excel_with_gain_summary_inline_India(
                original_file.read(),
                website_file.read(),
            )
            sheets = result.get("sheets", {})
            target = sheets.get(sheet_name)
            if not target:
                return JsonResponse({
                    "sheet_name": sheet_name,
                    "common_rows": [],
                    "different_rows": [],
                    "only_in_original_rows": [],
                    "only_in_website_rows": []
                })
            return JsonResponse({
                "sheet_name": sheet_name,
                "common_rows": target.get("common_rows", []),
                "different_rows": target.get("different_rows", []),
                "only_in_original_rows": target.get("only_in_original_rows", []),
                "only_in_website_rows": target.get("only_in_website_rows", [])
            })
        else:
            result = compare_single_sheet_diff_inline(
                original_file.read(),
                website_file.read(),
                sheet_name,
            )
            return JsonResponse(result)
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)


@csrf_exempt
def compare_excel_file_api(request):
    if request.method != "POST":
        return JsonResponse({"error": "Invalid request"}, status=400)
    try:
        original_file = request.FILES.get("original_file")
        website_file = request.FILES.get("website_file")
        sheets_config = request.POST.get("sheets_config")
        country = request.POST.get("country", "usa")

        parsed_config = None
        if sheets_config:
            try:
                parsed_config = json.loads(sheets_config)
            except Exception:
                parsed_config = None

        output = BytesIO()
        if country == "india":
            generate_comparison_excel_India(
                original_file.read(),
                website_file.read(),
                output,
                sheets_config=parsed_config,
            )
        else:
            generate_comparison_excel(
                original_file.read(),
                website_file.read(),
                output,
                sheets_config=parsed_config,
            )
        output.seek(0)

        response = FileResponse(output, as_attachment=True, filename="Comparison_Result.xlsx")
        return response
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)