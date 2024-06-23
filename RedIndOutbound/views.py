from django.http import HttpResponse, JsonResponse
from django.shortcuts import render
from rest_framework.decorators import api_view
from rest_framework.response import Response
from teams.models import Login


@api_view(['GET'])
def login(request):

    serialized_items = []
    usernames = Login.objects.all()

    for username in usernames:
        print(username)
        serialized_items.append({
            'Username': username.username,
        })

    print(serialized_items)

    return JsonResponse(serialized_items, safe=False)

@api_view(['POST'])
def processlogin(request):

    email = request.query_params.get("username")
    password = request.query_params.get("password")
    #logindetails = (request.data)
    #for logindetail in logindetails:
    #    email = logindetail["username"]
    #    password = logindetail["password"]
    #email = logindetail["username"]
    #passwor = logindetail["password"]

    login_query = Login.objects.all()
    loginbool = False

    for login in login_query:
        if str(login.username).upper()==str(email).upper() and str(login.password).upper()==str(password).upper():
            loginbool = True

    if loginbool == True and (str(email).upper()=="SATYANARAYANA.P@REDINGTONGROUP.COM" or str(email).upper()=="DEEPAK.KUMARU@REDINGTONGROUP.COM" or str(email).upper()=="UDHAYASHANKAR.S@REDINGTONGROUP.COM" ):
        return Response("Valid")
    elif loginbool == True:
        return Response("Valid")
    else:
        return Response("InValid")



