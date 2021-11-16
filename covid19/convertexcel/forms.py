from django.forms import ModelForm
from django import forms
from .models import *

class FileUploadForm(ModelForm):
    class Meta:
        model = FileUpload
        fields = ['excel_file']
        widgets = {
            'excel_file': forms.FileInput(
                attrs = {
                    'class': 'form-control form-control-lg mt-3'
                }
            )
        }
        labels = {
            'excel_file': ''
        }
        