"""Step 2 · Build the four HTML files in the project folder from template.html + prices.json.
python3 source/build.py"""
import json, os
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
src = lambda f: os.path.join(HERE, f); dst = lambda f: os.path.join(ROOT, f)
t = open(src('template.html')).read()
import base64
t = t.replace('/*__LOGO__*/', 'data:image/jpeg;base64,' + base64.b64encode(open(src('joulewise-logo-small.jpg'), 'rb').read()).decode())  # JouleWise logo, embedded so it works offline
r = json.load(open(src('prices.json'))); d = {'dates': r['dates'][-365:]}   # latest 365 days
for m in ['DAM', 'GDAM', 'RTM']: d[m] = r[m][-365:]
full = t.replace('/*__DATA__*/', json.dumps(d, separators=(',', ':')))
open(dst('bess_arbitrage_model.html'), 'w').write(full)          # artifact version (no <html> wrapper)
head = '<!doctype html>\n<html lang="en">\n<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
def standalone(title, tail):
    s = full.replace('<title>IEX BESS Arbitrage</title>', f'<title>{title}</title>', 1)
    s = head + s.replace('[hidden]{display:none!important}', '[hidden]{display:none!important}\nbody{margin:0}', 1)
    return s + '\n' + tail + '\n</html>\n'
full_sa = standalone('IEX BESS Arbitrage', '<!-- full two-tab model, standalone -->')
open(dst('bess_arbitrage_full_standalone.html'), 'w').write(full_sa)
open(dst('index.html'), 'w').write(full_sa)   # Vercel / static-host home page = full two-tab model
open(dst('solar_bess_iex_arbitrage.html'), 'w').write(standalone('Solar BESS Arbitrage',
    "<script>\n(function(){ document.querySelector('.viewtabs').hidden = true; document.getElementById('view2').remove(); VIEW = 1; })();\n</script>"))
open(dst('two_cycle_catl_iex.html'), 'w').write(standalone('CATL Two-Cycle Arbitrage', open(src('standalone.js')).read()))
s = full[full.rindex('<script>')+8:full.rindex('</script>')]
open(src('.model.js'), 'w').write(s)   # extracted script, used by export_dispatch.py and for `node --check`
print('built 5 HTML files (incl. index.html) in', ROOT)
