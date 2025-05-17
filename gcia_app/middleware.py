# your_app/middleware.py
from django.utils.timezone import now
from datetime import timedelta
from django.shortcuts import redirect

class SessionTimeoutMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        if request.user.is_authenticated:
            # Check the last activity time in the session
            last_activity = request.session.get('last_activity')
            if last_activity and now() - timedelta(minutes=30) > now().fromisoformat(last_activity):
                from django.contrib.auth import logout
                logout(request)
                return redirect('login')  # Redirect to the login page on timeout
            # Update last activity time
            request.session['last_activity'] = now().isoformat()

        return self.get_response(request)
    

class LoginRedirectMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        response = self.get_response(request)
        # If we hit a 404 for /accounts/login/, redirect to the correct login URL
        if response.status_code == 404 and request.path == '/accounts/login/':
            next_url = request.GET.get('next', '')
            if next_url:
                return redirect(f'/app/login/?next={next_url}')
            return redirect('/app/login/')
            
        return response
