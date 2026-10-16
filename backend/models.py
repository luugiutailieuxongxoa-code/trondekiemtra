from typing import Literal, Optional, Union
from pydantic import BaseModel, Field
Mode = Literal["g0","g1","g2","g3"]
class Option(BaseModel):
    id: str
    text: str
class Question(BaseModel):
    id: str
    number: int = 1
    type: Literal["multiple_choice","true_false","short_answer","essay"] = "multiple_choice"
    topic: str = "Chưa phân loại"
    level: Literal["Biết","Hiểu","Vận dụng"] = "Biết"
    text: str
    options: list[Option] = Field(default_factory=list)
    answer: Optional[Union[str,list[str]]] = None
    points: float = 0.25
    group: Mode = "g0"
class MixRequest(BaseModel):
    questions: list[Question]
    mode: Mode = "g3"
    code_count: int = Field(default=4, ge=1, le=50)
    question_count: Optional[int] = Field(default=None, ge=1)
    seed: Optional[int] = None
