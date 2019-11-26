from django import forms

class locationForm(forms.Form):
    file = forms.FileField()