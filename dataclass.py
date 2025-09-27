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
    step: Optional[str] = None


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


@dataclass
class searchResult:
    sheet_name: str
    unfinished: int = field(default=0)
    total: int = field(default=0)
    results: list[patternInfo] = field(default_factory=list)
    max_col_used: int = field(default=9)


@dataclass
class RegisteredDocumentAttachment:
    attachment_id: str          # XML 属性 attachmentId
    document_no: str            # <DocumentNo>
    file_name: str              # <FileName>


@dataclass
class Recipient:
    name: str
    organization_name: str


@dataclass
class FromUserDetails:
    name: str
    organization_name: str


@dataclass
class MailDetail:
    mail_id: str
    subject: str
    sent_date: datetime
    mail_data: str

    from_user_details: FromUserDetails
    attachments: list[RegisteredDocumentAttachment] = field(default_factory=list)
    recipients: list[Recipient] = field(default_factory=list)
