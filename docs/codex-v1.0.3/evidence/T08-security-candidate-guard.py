
import contextlib, hashlib, io, json, pathlib, platform, sys
wheel=pathlib.Path(sys.argv[1]); sys.path.insert(0,str(wheel))
import pypdf
from pypdf import apply_configuration, get_configuration
from pypdf.errors import LimitReachedError
from pypdf.generic import ArrayObject, DictionaryObject, ContentStream, DecodedStreamObject

def parse_content(data):
    stream=DecodedStreamObject(); stream.set_data(data)
    return ContentStream(stream,None).operations
cases=[('array','array_maximum_element_count',lambda data:ArrayObject.read_from_stream(io.BytesIO(data),None),b'[0 1 2]',b'[0 1 2 3]'),('dictionary','dictionary_maximum_element_count',lambda data:DictionaryObject.read_from_stream(io.BytesIO(data),None),b'<</A 0 /B 1 /C 2>>',b'<</A 0 /B 1 /C 2 /D 3>>'),('content','stream_parsing_work_maximum_count',parse_content,b'q Q q',b'q Q q Q')]
rows=[]
for name,field,parser,valid,over in cases:
    exists=hasattr(get_configuration(),field)
    with apply_configuration(**{field:3}) if exists else contextlib.nullcontext():
        positive=parser(valid)
        try:
            value=parser(over)
        except LimitReachedError as error:
            raised=True; message=str(error)
        else:
            raised=False; message='accepted '+str(len(value))+' elements/operations'
    rows.append({'case':name,'configuration_field':field,'default':getattr(get_configuration(),field,None),'lowered_limit':3 if exists else None,'positive_control_length':len(positive),'over_limit_rejected':raised,'message':message,'restored_default':getattr(get_configuration(),field,None)})
expected=pypdf.__version__=='6.21.1'
assert all(row['positive_control_length']==3 and row['over_limit_rejected']==expected for row in rows)
print(json.dumps({'version':pypdf.__version__,'developer_python':platform.python_version(),'import_origin':pypdf.__file__,'wheel_sha256':hashlib.sha256(wheel.read_bytes()).hexdigest(),'result':'CANDIDATE_GUARDS_PASS' if expected else 'OLD_VERSION_GUARDS_ABSENT_REPRODUCED','cases':rows}))
