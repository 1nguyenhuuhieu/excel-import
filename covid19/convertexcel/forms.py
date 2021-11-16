from django.forms import ModelForm
from .models import *

class FileUploadForm(ModelForm):
    class Meta:
        model = FileUpload
        fields = ['excel_file']
        