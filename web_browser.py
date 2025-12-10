import webbrowser

for i in range(155, 161):
    url = f'http://192.168.4.{i}/'
    webbrowser.open_new_tab(url)