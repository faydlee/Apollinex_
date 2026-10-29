from django import forms
from django.contrib.auth import get_user_model


class AccountProfileForm(forms.ModelForm):
    """Edit only the display name fields on the current user's account."""

    class Meta:
        model = get_user_model()
        fields = ("first_name", "last_name")
        labels = {
            "first_name": "Имя",
            "last_name": "Фамилия",
        }
