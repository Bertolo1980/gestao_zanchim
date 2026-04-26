from django.conf import settings
from django.shortcuts import redirect


class LoginRequiredForInternalPagesMiddleware:
    PUBLIC_PATHS = {
        '/',
        '/accounts/login/',
    }

    PUBLIC_PREFIXES = (
        settings.STATIC_URL,
        settings.MEDIA_URL,
        '/admin/',
        '/webhook/',
    )

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        path = request.path
        if (
            path not in self.PUBLIC_PATHS
            and not any(path.startswith(prefix) for prefix in self.PUBLIC_PREFIXES if prefix)
            and not request.user.is_authenticated
        ):
            return redirect(f"{settings.LOGIN_URL}?next={path}")
        return self.get_response(request)
