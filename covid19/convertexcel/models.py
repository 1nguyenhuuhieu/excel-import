from django.db import models
import uuid

# Create your models here.

class FileUpload(models.Model):
    id = models.UUIDField(
         primary_key = True,
         default = uuid.uuid4,
         editable = False)
    excel_file = models.FileField(upload_to="static/excel/")
    output = models.URLField( blank=True)
