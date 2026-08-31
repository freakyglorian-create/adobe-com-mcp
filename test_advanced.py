import subprocess, json, threading, time, os

SERVER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'server.py')
proc = subprocess.Popen(['python', SERVER],
    stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
    text=True, encoding='utf-8', bufsize=1)

responses = {}
def reader():
    for line in proc.stdout:
        line = line.strip()
        if not line: continue
        try:
            d = json.loads(line)
            if 'id' in d: responses[d['id']] = d
        except: pass

t = threading.Thread(target=reader, daemon=True)
t.start()
def send(msg):
    proc.stdin.write(json.dumps(msg, ensure_ascii=False) + '\n')
    proc.stdin.flush()
def wait(rid, timeout=20):
    s = time.time()
    while time.time() - s < timeout:
        if rid in responses: return responses[rid]
        time.sleep(0.1)
    return None
def call(rid, name, args=None):
    if args is None: args = {}
    send({'jsonrpc':'2.0','id':rid,'method':'tools/call','params':{'name':name,'arguments':args}})
    r = wait(rid, 25)
    if r and 'result' in r:
        for c in r['result'].get('content',[]):
            if c.get('type') == 'text':
                try: return json.loads(c['text'])
                except: return c['text'][:300]
    elif r and 'error' in r:
        return "ERROR: " + str(r.get('error'))
    return 'TIMEOUT'

time.sleep(2)
send({'jsonrpc':'2.0','id':1,'method':'initialize','params':{'protocolVersion':'2024-11-05','capabilities':{},'clientInfo':{'name':'t','version':'1.0'}}})
wait(1, 10)
send({'jsonrpc':'2.0','method':'notifications/initialized'})
time.sleep(0.3)

print('=== 1. Create test doc 600x400 ===')
print(call(2, 'ps_create_document', {'width':600,'height':400,'name':'AdvancedTest'}))

print('\n=== 2. Add text layer 1 ===')
print(call(3, 'ps_add_text_layer', {'text':'Hello World','size':36,'x':50,'y':50,'r':255,'g':100,'b':50}))

print('\n=== 3. Add text layer 2 ===')
print(call(4, 'ps_add_text_layer', {'text':'Second Text','size':24,'x':50,'y':150,'r':50,'g':100,'b':255}))

print('\n=== 4. Fill background ===')
print(call(5, 'ps_fill_layer', {'r':40,'g':40,'b':60}))

print('\n=== 5. List text layers ===')
print(call(6, 'ps_list_text_layers'))

print('\n=== 6. Select all text layers ===')
print(call(7, 'ps_select_all_text_layers'))

print('\n=== 7. List all layers ===')
print(call(9, 'ps_list_all_layers'))

print('\n=== 8. Extract color palette ===')
print(call(10, 'ps_extract_color_palette', {'num_colors':5}))

proc.terminate()
proc.wait(timeout=3)
