import zipfile, sys
sys.stdout.reconfigure(encoding='utf-8')

def get_rels_xml(path):
    with zipfile.ZipFile(path) as z:
        return z.read('word/_rels/document.xml.rels').decode('utf-8')

orig_xml  = get_rels_xml('02 ВКР Али.docx')
final_xml = get_rels_xml('02_VKR_Ali_final.docx')

# Show raw snippet to debug quote style
print("ORIG  first 600 chars:")
print(repr(orig_xml[:600]))
print()
print("FINAL first 600 chars:")
print(repr(final_xml[:600]))

# Count image relationships
orig_img  = orig_xml.count('/relationships/image')
final_img = final_xml.count('/relationships/image')
print(f"\nImage rels — orig: {orig_img}, final: {final_img}")

# Check if all image rIds are present in both
import re
def get_image_rids(xml):
    return set(re.findall(r'Id=.([^"\']+).[^>]+image', xml))

orig_rids  = get_image_rids(orig_xml)
final_rids = get_image_rids(final_xml)
print(f"Image rIds — orig: {len(orig_rids)}, final: {len(final_rids)}")
print(f"Only in orig:  {orig_rids - final_rids}")
print(f"Only in final: {final_rids - orig_rids}")
