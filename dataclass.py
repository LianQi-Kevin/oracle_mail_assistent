from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional


@dataclass
class responseMailInfo:
    mailID: int
    MailNo: str
    SentDate: datetime
    subject: str
    AllAttachmentCount: int


@dataclass
class patternInfo:
    unit: str
    discipline: str
    drawing: str
    wf: Optional[str] = None
    ver: Optional[str] = "*"
    title: Optional[str] = None


@dataclass
class UserRef:
    organization_id: int
    organization_name: str
    name: str
    user_id: int


@dataclass
class Workflow:
    workflow_id: int
    step_name: str
    step_outcome: str
    step_status: str
    date_in: datetime
    date_completed: Optional[datetime]
    date_due: datetime
    days_late: int
    duration: float

    document_number: str
    document_revision: str
    document_title: str
    document_version: int
    file_name: str
    file_size: int

    initiator: UserRef
    reviewer: Optional[UserRef]
    assignees: list[UserRef] = field(default_factory=list)


@dataclass
class WorkflowSearchResult:
    current_page: int
    page_size: int
    total_pages: int
    total_results: int
    total_results_on_page: int
    workflows: list[Workflow]
