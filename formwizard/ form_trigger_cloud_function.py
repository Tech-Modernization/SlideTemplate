# this just echos out the form request json object
# TODO: create a Drive folder for project, copy some template files there, apply templating
def form_trigger(request):
    payload = request.get_json(force=True, silent=True)
    print(f"Payload was: {payload}")
    return "OK"