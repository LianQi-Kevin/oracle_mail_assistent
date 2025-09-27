"""example"""

from main import requestToken, searchMail, MAIN_RE, searchWorkflow
from dataclass import patternInfo

if __name__ == '__main__':
    requestToken()

    data = ["SLDS-BCEG-002-SDS-I-I064", "SLDS-BCEG-002-SDS-I-I065", "SLDS-BCEG-002-SDS-I-I066",
            "SLDS-BCEG-002-SDS-I-I067", "SLDS-BCEG-002-SDS-I-I068", "SLDS-BCEG-002-SDS-I-I069",
            "SLDS-BCEG-001-SDS-I-I001", "SLDS-BCEG-001-0405-SDS-I-I001", "SLDS-BCEG-001-0405-SDS-I-I002"]

    for ID in data:
        matched_data = MAIN_RE.match(ID).groupdict()
        response = searchMail(patternInfo(
                unit=matched_data["unit"], discipline=matched_data["discipline"],
                drawing=matched_data["drawing"], step=matched_data["step"]
            ))[0]
        matched_subject = MAIN_RE.match(response.subject).groupdict()
        print(f"matched subject: {matched_subject}")
        print(f"sentDate: {response.SentDate.date().isoformat()}")
        if matched_subject.get('wf'):
            wf_response = searchWorkflow(workflow_num=matched_subject.get('wf'))
            for workflow in wf_response.workflows:
                print(f"name: {workflow.assignees[0].name}, step out name: {workflow.step_outcome}")
