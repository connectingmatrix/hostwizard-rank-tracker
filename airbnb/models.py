from django.db import models
from datetime import datetime


class GoogleSheet(models.Model):
    sheet_name = models.CharField(max_length=200, default='Automation Ranking')
    sheet_url = models.CharField(max_length=200)
    created_at = models.DateTimeField('created_at', default=datetime.now)

    def __str__(self):
        return self.sheet_name
