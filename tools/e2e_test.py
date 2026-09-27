import json
from playwright.sync_api import sync_playwright
import pathlib; out=str(pathlib.Path(__file__).parent)+'/'
got=[]
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':820,'height':1180})
    errs=[]; pg.on('pageerror',lambda e:errs.append(str(e)))
    def handle(route):
        got.append(json.loads(route.request.post_data)); route.fulfill(status=200,body='{"ok":true}',headers={'content-type':'application/json','access-control-allow-origin':'*'})
    pg.route('https://script.google.com/**',handle)
    pg.goto(pathlib.Path(__file__).resolve().parents[1].joinpath('index.html').as_uri())
    pg.evaluate("localStorage.setItem('wl_settings',JSON.stringify({url:'https://script.google.com/macros/s/X/exec',key:'pechanga-fc-waste',showCost:false}))")
    pg.reload()
    pg.fill('#chef','Test Chef'); pg.fill('#badge','12345'); pg.click('#startBtn')
    names=['Pechanga Fried Chicken','Pronto','Little Wok','Agave']
    for k,n in enumerate(names):
        pg.click(f'.outlet:has-text("{n}")')
        items=pg.locator('#countList .item')
        for i in range(k+2):
            items.nth(i).locator('input').fill(str(3*(i+1))); items.nth(i).locator('input').press('Tab')
        if k==0:
            pg.click('#finishOutletBtn')  # should block: blanks
            pg.screenshot(path=out+'b_block.png')
            print('still count view:', pg.is_visible('#vCount'))
        pg.click('#zeroRestBtn'); pg.fill('#notes', f'note {n}')
        pg.click('#finishOutletBtn')
        if k==1:
            print('review disabled after 2:', pg.is_disabled('#hubReviewBtn'))
            pg.screenshot(path=out+'b_hub.png')
            pg.reload(); print('resumed hub:', pg.is_visible('#vHub'))
    print('review disabled after 4:', pg.is_disabled('#hubReviewBtn'))
    pg.click('#hubReviewBtn'); pg.screenshot(path=out+'b_review.png',full_page=True)
    pg.click('#submitBtn'); pg.wait_for_selector('#vDone:not(.hidden)')
    print('errs',errs, 'posts',len(got))
    g=got[0]; print(g['totals'], [(s['stationLabel'],s['totals'],s['closeTime']) for s in g['stations']])
    json.dump(g,open(out+'payload2.json','w'))
