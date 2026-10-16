import random
from copy import deepcopy
def mix_exam(req):
    rng=random.Random(req.seed)
    source=[q.model_dump() for q in req.questions]
    if req.question_count: source=source[:req.question_count]
    codes=[]
    for n in range(req.code_count):
        qs=deepcopy(source)
        if req.mode in ("g1","g3"): rng.shuffle(qs)
        for q in qs:
            if req.mode in ("g2","g3") and q["type"]=="multiple_choice": rng.shuffle(q["options"])
        for i,q in enumerate(qs,1): q["number"]=i
        codes.append({"code":str(n+1).zfill(3),"questions":qs})
    warnings=["Đáp án được gắn với ID phương án, không theo chữ A/B/C/D.","Tiểu ý đúng/sai giữ nguyên thứ tự trong MVP.","Cần giáo viên xác minh đáp án và đối chiếu công thức/hình ảnh với file gốc."]
    return {"codes":codes,"warnings":warnings}
