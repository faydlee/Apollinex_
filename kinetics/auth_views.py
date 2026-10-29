from django.contrib import messages
from django.contrib.auth import login, update_session_auth_hash
from django.contrib.auth.decorators import login_not_required, login_required
from django.contrib.auth.forms import PasswordChangeForm, UserCreationForm
from django.contrib.auth.views import LoginView as DjangoLoginView
from django.contrib.auth.views import LogoutView as DjangoLogoutView
from django.shortcuts import redirect, render
from django.urls import reverse_lazy
from django.utils.decorators import method_decorator
from django.views.generic import FormView

from .auth_forms import AccountProfileForm


def _password_change_form(user, data=None):
    form = PasswordChangeForm(user, data=data)
    form.fields["old_password"].label = "Текущий пароль"
    form.fields["new_password1"].label = "Новый пароль"
    form.fields["new_password2"].label = "Подтверждение нового пароля"
    form.fields["old_password"].widget.attrs["autocomplete"] = "current-password"
    form.fields["new_password1"].widget.attrs["autocomplete"] = "new-password"
    form.fields["new_password2"].widget.attrs["autocomplete"] = "new-password"
    return form


@login_required
def settings_page(request):
    """Update the signed-in user's profile and password using Django forms."""

    profile_form = AccountProfileForm(instance=request.user)
    password_form = _password_change_form(request.user)

    if request.method == "POST":
        section = request.POST.get("section")
        if section == "profile":
            profile_form = AccountProfileForm(
                request.POST,
                instance=request.user,
            )
            if profile_form.is_valid():
                profile_form.save()
                messages.success(request, "Изменения учётной записи сохранены.")
                return redirect("settings")
            messages.error(request, "Проверьте поля учётной записи.")
        elif section == "password":
            password_form = _password_change_form(request.user, request.POST)
            if password_form.is_valid():
                user = password_form.save()
                update_session_auth_hash(request, user)
                messages.success(request, "Пароль успешно изменён.")
                return redirect("settings")
            messages.error(request, "Не удалось изменить пароль. Проверьте форму.")
        else:
            messages.error(request, "Неизвестный раздел настроек.")

    return render(
        request,
        "kinetics/account_settings.html",
        {
            "profile_form": profile_form,
            "password_form": password_form,
        },
    )


@method_decorator(login_not_required, name="dispatch")
class LoginView(DjangoLoginView):
    template_name = "kinetics/auth_login.html"
    redirect_authenticated_user = True

    def form_valid(self, form):
        response = super().form_valid(form)
        messages.success(
            self.request,
            f"Вы вошли как {form.get_user().get_username()}.",
        )
        return response


@method_decorator(login_not_required, name="dispatch")
class RegistrationView(FormView):
    template_name = "kinetics/auth_register.html"
    form_class = UserCreationForm
    success_url = reverse_lazy("home")

    def dispatch(self, request, *args, **kwargs):
        if request.user.is_authenticated:
            return redirect("home")
        return super().dispatch(request, *args, **kwargs)

    def form_valid(self, form):
        user = form.save()
        login(self.request, user)
        messages.success(
            self.request,
            f"Учётная запись создана. Вы вошли как {user.get_username()}.",
        )
        return super().form_valid(form)

    def form_invalid(self, form):
        messages.error(
            self.request,
            "Не удалось создать учётную запись. Проверьте ошибки в форме.",
        )
        return super().form_invalid(form)


class LogoutView(DjangoLogoutView):
    next_page = reverse_lazy("home")

    def post(self, request, *args, **kwargs):
        response = super().post(request, *args, **kwargs)
        messages.success(request, "Вы вышли из учётной записи.")
        return response
