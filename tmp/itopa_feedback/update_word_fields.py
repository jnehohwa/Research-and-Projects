import sys
import time

import uno
from com.sun.star.beans import PropertyValue


def property_value(name, value):
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


document_path = sys.argv[1]
local_context = uno.getComponentContext()
resolver = local_context.ServiceManager.createInstanceWithContext(
    'com.sun.star.bridge.UnoUrlResolver', local_context
)

last_error = None
for _ in range(40):
    try:
        context = resolver.resolve(
            'uno:socket,host=127.0.0.1,port=20828;urp;StarOffice.ComponentContext'
        )
        break
    except Exception as error:
        last_error = error
        time.sleep(0.25)
else:
    raise RuntimeError(f'Could not connect to LibreOffice: {last_error}')

desktop = context.ServiceManager.createInstanceWithContext(
    'com.sun.star.frame.Desktop', context
)
document_url = uno.systemPathToFileUrl(document_path)
document = desktop.loadComponentFromURL(
    document_url,
    '_blank',
    0,
    (property_value('Hidden', True), property_value('UpdateDocMode', 3)),
)
if document is None:
    raise RuntimeError('LibreOffice could not open the document')

indexes = document.getDocumentIndexes()
for index_number in range(indexes.getCount()):
    indexes.getByIndex(index_number).update()

try:
    document.getTextFields().refresh()
except Exception:
    pass

document.calculateAll()
document.store()
document.close(True)
print(f'Updated fields in {document_path}')
