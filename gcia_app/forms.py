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
