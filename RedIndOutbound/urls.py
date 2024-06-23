from django.urls import path
from . import  views

app_name="redindoutbound"

urlpatterns = [
    path('login-users-email',views.login,name='login-users-email'),
    path('process-login',views.processlogin,name='process-login'),
]
