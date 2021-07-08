from django.shortcuts import render, redirect
from django.contrib.auth.models import auth
from django.contrib import messages
from .tasks import logic

# Create your views here.
def index(request):
    if request.method == 'POST':
        username = request.POST['username']
        password = request.POST['password']
        user = auth.authenticate(username = username, password = password)
        if user is not None:
            auth.login(request, user)
            return redirect(upload)
        else:
            messages.info(request, "Invalid Credentials")
            return redirect("/")
    return render(request, "index.html")

def upload(request):
    if request.method=='POST':
        if request.FILES['inputfile']==None:
            return render(request, "upload.html")
        else:
            file1 = request.FILES['inputfile']
            with open('rfrapp/input/input.xlsx', 'wb+') as destination:
                for chunk in file1.chunks():
                    destination.write(chunk)
            return redirect(rfrstatus)
    if request.user.is_authenticated:
        return render(request, "upload.html")

def rfrstatus(request):
    if request.user.is_authenticated:
        logic_task = logic.delay()
        task_id = logic_task.task_id
        return render(request, "status.html", {'task_id':task_id})

def logout(request):
    auth.logout(request)
    return redirect("/")