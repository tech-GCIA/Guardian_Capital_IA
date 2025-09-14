from django import forms
from django.contrib.auth.forms import UserCreationForm
from gcia_app.models import Customer

class CustomerCreationForm(UserCreationForm):
    class Meta:
        model = Customer
        fields = ['username', 'email', 'password1', 'password2']
        widgets = {
            'username': forms.TextInput(attrs={'placeholder': 'Enter your username'}),
            'email': forms.EmailInput(attrs={'placeholder': 'Enter your email address'}),
            # Here we are overriding the PasswordInput widget for both password fields
            'password1': forms.PasswordInput(attrs={'placeholder': 'Enter a secure password'}),
            'password2': forms.PasswordInput(attrs={'placeholder': 'Confirm your password'}),
        }
        help_texts = {
            'username': None,  # Removes help text for username
            'email': None,     # Just in case email has help text
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Explicitly remove help texts for password fields
        self.fields['password1'].help_text = None
        self.fields['password2'].help_text = None
        
        # Explicitly add the placeholder text to the password fields if they aren't rendered correctly
        self.fields['password1'].widget.attrs.update({'placeholder': 'Enter a secure password'})
        self.fields['password2'].widget.attrs.update({'placeholder': 'Confirm your password'})

class ExcelUploadForm(forms.Form):
    excel_file = forms.FileField()

class MasterDataExcelUploadForm(forms.Form):
    """
    Form for uploading Excel files with a dropdown to select file type
    """
    file_type = forms.ChoiceField(
        choices=[
            ('top_schemes', 'Top Schemes Data'),
            ('ratios_pe', 'Ratios, PE Data'),
            ('index_nav', 'INDEX NAV')
        ],
        widget=forms.Select(attrs={'class': 'form-control'}),
        label="File Type"
    )
    excel_file = forms.FileField(
        widget=forms.FileInput(attrs={'class': 'form-control-file'}),
        label="Excel File"
    )

# Add this to gcia_app/forms.py

from django import forms
from django.core.exceptions import ValidationError
import os

class StockDataUploadForm(forms.Form):
    """
    Form for uploading Excel files containing stock data
    """
    excel_file = forms.FileField(
        widget=forms.FileInput(attrs={
            'class': 'form-control-file',
            'accept': '.xls,.xlsx,.xlsm',
            'id': 'stockDataFile'
        }),
        label="Stock Data Excel File",
        help_text="Upload an Excel file (.xls, .xlsx, .xlsm) containing stock data from App-Base Sheet"
    )
    
    def clean_excel_file(self):
        """
        Validate the uploaded Excel file
        """
        excel_file = self.cleaned_data['excel_file']
        
        # Check file extension
        if excel_file:
            file_extension = os.path.splitext(excel_file.name)[1].lower()
            allowed_extensions = ['.xls', '.xlsx', '.xlsm']
            
            if file_extension not in allowed_extensions:
                raise ValidationError(
                    f"Invalid file type. Please upload an Excel file with one of these extensions: {', '.join(allowed_extensions)}"
                )
            
            # Check file size (limit to 50MB)
            if excel_file.size > 50 * 1024 * 1024:  # 50MB in bytes
                raise ValidationError("File size too large. Please upload a file smaller than 50MB.")
            
            # Check if file has content
            if excel_file.size == 0:
                raise ValidationError("The uploaded file is empty. Please upload a valid Excel file.")
        
        return excel_file
    

