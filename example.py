"""example"""

from main import requestToken, searchMail, MAIN_RE, searchWorkflow
from dataclass import patternInfo

if __name__ == '__main__':
    requestToken()

    data = ["SLDS-BCEG-002-SDS-I-I064", "SLDS-BCEG-002-SDS-I-I065", "SLDS-BCEG-002-SDS-I-I066",
            "SLDS-BCEG-002-SDS-I-I067", "SLDS-BCEG-002-SDS-I-I068", "SLDS-BCEG-002-SDS-I-I069"]

    for ID in data:
        try:
            matched = MAIN_RE.match(ID).groupdict()
            response = searchMail(
                patternInfo(unit=matched.get('unit', '002'), discipline=matched.get('discipline', 'I'),
                            drawing=matched.get('drawing', '')))[0]
            matched_subject = MAIN_RE.match(response.subject).groupdict()
            print(f"matched subject: {matched_subject}")
            print(f"sentDate: {response.SentDate.date().isoformat()}")

            wf_response = searchWorkflow(workflow_num=matched_subject.get('wf'))
            for workflow in wf_response.workflows:
                print(f"name: {workflow.assignees[0].name}, step out name: {workflow.step_outcome}")

        except Exception:
            pass
