from base64 import b64encode
from pathlib import Path

root = Path(__file__).parent
photos = [f"_thinkfast-photo-{index:02d}.png" for index in range(1, 18)]
videos = ["think-fast-video-01.mp4", "think-fast-video-02.mp4"]
assets = {}
for filename in photos + videos:
    path = root / filename
    mime = "video/mp4" if path.suffix.lower() == ".mp4" else "image/png"
    assets[filename] = f"data:{mime};base64,{b64encode(path.read_bytes()).decode('ascii')}"

mapping = {
    1: [photos[0]], 2: photos[1:4], 3: photos[4:6] + [videos[1]],
    4: photos[6:8], 5: [photos[8]], 6: [photos[9]], 7: photos[10:12],
    8: photos[12:14], 9: [photos[14], videos[0]], 10: photos[15:17],
}
groups = [{"group": str(group), "media": items} for group, items in mapping.items()]

page = r'''<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Think Fast review media import</title><style>
:root{color-scheme:dark;--gold:#e5b842}*{box-sizing:border-box}body{margin:0;padding:24px;background:#0d0d0c;color:#f3ead8;font:15px/1.5 Arial,sans-serif}.wrap{max-width:820px;margin:auto}h1{color:var(--gold)}p{color:#c9c0b1}.field{display:grid;gap:8px;margin:24px 0}input{width:100%;padding:12px;border:1px solid #68562b;border-radius:8px;background:#17130f;color:white}button{padding:13px 18px;border:0;border-radius:8px;background:var(--gold);color:#17130f;font-weight:700;cursor:pointer}button:disabled{opacity:.55;cursor:wait}.groups{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:8px}.group{padding:12px;border:1px solid #45391e;border-radius:8px;background:#17130f}.status{min-height:3em;margin-top:18px;color:var(--gold);white-space:pre-wrap}
</style></head><body><main class="wrap"><h1>Think Fast — review media</h1>
<p>Prepared as 10 separate Draft rows. Customer names, phone numbers, ratings, and comments are left blank. Group 3 includes video 2; Group 9 includes video 1.</p>
<div class="groups" id="groups"></div>
<label class="field">Google Apps Script Web App URL<input id="scriptUrl" type="url" placeholder="https://script.google.com/macros/s/.../exec" autocomplete="url"></label>
<button id="upload">Upload 10 Draft rows to Reviews</button><div class="status" id="status" role="status"></div>
<p>Before uploading, deploy the updated <code>code.gs</code> as a Web App. This page sends the prepared photos and videos to the script's Drive folder and spreadsheet. Check the Reviews tab after the upload finishes.</p>
</main><script>
const assets = __ASSETS__;
const groups = __GROUPS__;
document.getElementById('groups').innerHTML = groups.map(group => `<div class="group"><b>Group ${String(group.group).padStart(2,'0')}</b><br>${group.media.filter(name=>name.endsWith('.png')).length} photos · ${group.media.filter(name=>name.endsWith('.mp4')).length} videos</div>`).join('');
const button = document.getElementById('upload');
button.addEventListener('click', async () => {
  const url = document.getElementById('scriptUrl').value.trim();
  const status = document.getElementById('status');
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec(?:\?.*)?$/.test(url)) { status.textContent = 'Paste the deployed Google Apps Script Web App URL first.'; return; }
  if (!confirm('Send 17 customer photos and 2 videos to your Apps Script Drive folder and create/update 10 Draft rows in the Reviews sheet?')) return;
  button.disabled = true;
  status.textContent = 'Uploading media… Keep this page open until the request finishes.';
  try {
    const payload = { action: 'thinkFastBatchDrafts', groups: groups.map(group => ({ group: group.group, media: group.media.map(name => assets[name]) })) };
    await fetch(url, { method: 'POST', mode: 'no-cors', body: JSON.stringify(payload) });
    status.textContent = 'Upload request sent. Since Apps Script no-cors responses cannot be read here, open the Reviews sheet and confirm Groups 1–10 are present before closing this page.';
  } catch (error) {
    status.textContent = 'Upload failed: ' + (error.message || error);
    button.disabled = false;
  }
});
</script></body></html>'''

page = page.replace("__ASSETS__", __import__("json").dumps(assets, separators=(",", ":")))
page = page.replace("__GROUPS__", __import__("json").dumps(groups, separators=(",", ":")))
(root / "think-fast-review-import.html").write_text(page, encoding="utf-8")
print(f"Generated {root / 'think-fast-review-import.html'} ({(root / 'think-fast-review-import.html').stat().st_size:,} bytes)")
