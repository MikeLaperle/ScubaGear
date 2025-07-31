---
title: Microsoft Entra ID Collaboration Enablement Supplemental Baseline
version: MS.COLLAB.v1
author: Mike [submitter]
scope: Teams, Exchange, Entra ID
createdOn: 2025-07-31
reviewedOn: 2025-07-31
description: Extends Entra ID baseline to support secure collaboration across federal and DOD tenants with conditional access alignment.
---

# Collaboration Enablement Supplemental Baseline — `MS.COLLAB.v1`

## MS.COLLAB.1.1v1 — Teams Chat Federation with .gov / .mil

Enable secure Teams federation to allow external chat with trusted U.S. government and military domains.

**Rationale:** Supports interagency coordination using Microsoft Teams while maintaining control boundaries.

**NIST Mapping:** SC-7, AC-17(2)  
**MITRE ATT&CK:**  
- [T1219 – Remote Access Tools](https://attack.mitre.org/techniques/T1219/)  
- [T1071.001 – Web Protocols](https://attack.mitre.org/techniques/T1071/001/)

**Instructions (UI):**
- Navigate to **Teams admin center → External access**
- Toggle ON: _Users can communicate with other Teams users_
- Under **Allowed domains**, add `*.gov` and `*.mil`

---

## MS.COLLAB.1.2v1 — Exchange Free/Busy Calendar Sharing with .gov / .mil

Enable federated users to view free/busy availability across trusted government domains.

**Rationale:** Supports efficient scheduling across agencies while preserving privacy of calendar details.

**NIST Mapping:** AC-3, SC-12  
**MITRE ATT&CK:**  
- [T1114.002 – Remote Email Collection](https://attack.mitre.org/techniques/T1114/002/)  
- [T1087.002 – Domain Account Discovery](https://attack.mitre.org/techniques/T1087/002/)

**Instructions (UI):**
- Navigate to **Microsoft 365 admin center → Settings → Org Settings → Calendar**
- Enable _Calendar availability sharing_
- Choose: _Share free/busy with time only_
- Add trusted domains: `*.gov` and `*.mil`

---

## MS.COLLAB.1.3v1 — Cross-Tenant Access Default Settings

Configure default trust to allow secure collaboration with compliant identities across federal tenants.

**Rationale:** Promotes interoperability while enforcing strong authentication and device posture.

**NIST Mapping:** AC-21, SC-28  
**MITRE ATT&CK:**  
- [T1134.002 – Create Process with Token](https://attack.mitre.org/techniques/T1134/002/)  
- [T1078.004 – Valid Cloud Accounts](https://attack.mitre.org/techniques/T1078/004/)

**Instructions (UI):**
- Entra admin center → External Identities → Cross-tenant access settings → Default Settings
- Enable trust for MFA, compliant devices, hybrid join
- Inbound & Outbound B2B Collaboration → _Allow_

---

## MS.COLLAB.1.4v1 — Cross-Cloud Enablement for Entra B2B

Explicitly authorize B2B collaboration across commercial and government clouds.

**Rationale:** Supports hybrid-cloud agencies (GCC, GCC High) in secure cross-cloud exchanges.

**NIST Mapping:** AC-21, CA-3, AU-2  
**MITRE ATT&CK:**  
- [T1484.002 – Trust Modification](https://attack.mitre.org/techniques/T1484/002/)  
- [T1078.004 – Cloud Accounts](https://attack.mitre.org/techniques/T1078/004/)

**Instructions (UI):**
- Entra admin center → External Identities → Organizational settings
- Add partner tenant ID
- Set inbound/outbound collaboration → _Allow_
- Trust settings: MFA, compliant device, hybrid join
- Toggle ON: _Enable access between Microsoft Azure commercial and Azure Government clouds_

---

## MS.COLLAB.1.5v1 — External Collaboration Guest Access Restriction

Minimize directory metadata exposure for guest accounts.

**Rationale:** Reduces reconnaissance vectors for external actors.

**NIST Mapping:** AC-2(5), AC-3  
**MITRE ATT&CK:**  
- [T1087.001 – Account Discovery](https://attack.mitre.org/techniques/T1087/001/)  
- [T1069.002 – Domain Groups](https://attack.mitre.org/techniques/T1069/002/)

**Instructions (UI):**
- Entra admin center → External Collaboration Settings
- Select: _Guest users have limited access to properties and memberships of directory objects_

---

## MS.COLLAB.1.6v1 — External Collaboration Domain Restrictions

Restrict guest invitations to trusted government and military domains.

**Rationale:** Avoids unauthorized collaboration invitations and domain spoofing.

**NIST Mapping:** SC-7, CA-3  
**MITRE ATT&CK:**  
- [T1589.002 – Email Addresses](https://attack.mitre.org/techniques/T1589/002/)  
- [T1190 – Public-Facing Application Exploit](https://attack.mitre.org/techniques/T1190/)

**Instructions (UI):**
- Under External Collaboration Settings → Collaboration Restrictions
- Choose: _Allow invitations only to specified domains_
- Add: `*.gov`, `*.mil`, known partner domains

---

## MS.COLLAB.1.7v1 — Exclusion from Certificate-Based Authentication

Prevent external identities from using certificate-based authentication methods.

**Rationale:** Protects internal PKI systems and authentication integrity.

**NIST Mapping:** IA-2, IA-5  
**MITRE ATT&CK:**  
- [T1550.003 – Pass the Ticket](https://attack.mitre.org/techniques/T1550/003/)  
- [T1606.001 – Forge Web Credentials](https://attack.mitre.org/techniques/T1606/001/)

**Instructions (UI):**
- Entra admin center → Authentication Methods → CBA Policy
- Target: internal users only
- Exclude via domain match (see internal group definition)

---

## MS.COLLAB.1.8v1 — Dynamic Group for Internal Users (Domain Match)

Segment internal users using verified domain suffixes.

**Rationale:** Enhances accuracy beyond userType classifications.

**NIST Mapping:** AC-2, AC-5  
**MITRE ATT&CK:**  
- [T1078.004 – Cloud Accounts](https://attack.mitre.org/techniques/T1078/004/)  
- [T1484.002 – Trust Modification](https://attack.mitre.org/techniques/T1484/002/)

**Instructions (UI):**
- Groups → New → Security → Dynamic
- Rule:
  ```plaintext
  (user.mail -match "@agency.gov") -or (user.mail -match "@service.mil")
  ```
- Name: _Internal Users (Domain-Matched)_

---

## MS.COLLAB.1.9v1 — Dynamic Group for External Users (Domain Exclusion)

Segment external identities by excluding known internal domains.

**Rationale:** Prevents inaccurate scoping caused by guest promotions or hybrid identities.

**NIST Mapping:** AC-2(5), PE-2  
**MITRE ATT&CK:**  
- [T1087.001 – Account Discovery](https://attack.mitre.org/techniques/T1087/001/)  
- [T1098.001 – Additional Cloud Credentials](https://attack.mitre.org/techniques/T1098/001/)

**Instructions (UI):**
- Groups → New → Security → Dynamic
- Rule:
  ```plaintext
  -not ((user.mail -match "@agency.gov") -or (user.mail -match "@service.mil"))
  ```
- Name: _External Users (Domain-Excluded)_

---

*Note: Graph API automation and domain validation logic will be defined in `MS.COLLAB.v2`.*
