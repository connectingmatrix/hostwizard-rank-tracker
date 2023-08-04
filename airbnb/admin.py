from django.contrib import admin
from .models import GoogleSheet


class GoogleSheetAdmin(admin.ModelAdmin):
    fieldsets = [
        ("name", {'fields': ["sheet_name"]}),
        ("url", {"fields": ["sheet_url"]}),
        ("date", {"fields": ["created_at"]})
    ]


admin.site.register(GoogleSheet, GoogleSheetAdmin)
