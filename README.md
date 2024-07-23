
# About KNOCKOUT

The tool KNOCKOUT has been developed as a proof of concept tool which has been developed in combination with the blog post published on the NVISO blog: [From Evidence to Advantage: Leveraging Incident Response Artifacts for Red Team Engagements](BLOG-LINK).

Using KNOCKOUT now gives us the opportunity to collect multiple sources of information in one step to efficiently assess our situation on the compromised endpoint. This otherwise would require multiple steps to query the different sources (file system, registry, etc.), then aggregate the information from the various output formats which as a result is taking a chunk of time which is already very valuable and limited during initial access, where every minute counts to proceed with persistence or lateral movement to avoid loosing the foothold.

# Features

The following table provides an overview of what information the developed tool collects and what we as red team operators can learn from the collected information.

|Artifact|Learning|
|---|---|
|Recently used Office files and folders|Has our victim worked on documents related to our target systems or sensitive information that can be used to achieve an objective?|
|Shortcuts (LNK or URL) in common locations|Are commonly used work folders on network shares or cloud applications linked in common locations?
This gives us an insight on commonly used internal portals for IT service desks, intranet pages, code repositories as well as if network profiles or other cloud syncing solutions are in use.|
|Explorer files and folders|What are recently accessed files and their locations? What are executed commands using the Run dialog? What paths have been manually entered paths in Windows Explorer?|
|History of connected USB storage devices|Are mobile storage devices allowed or can they be used for data exchange?
This would allow for a malicious USB stick scenario or data exfiltration using USB storage devices.|
|Browser Favorites (Microsoft Edge for now)|Browser favorites can give us an insight on commonly used internal portals for IT service desks, intranet pages, code repositories and more.|
|Usage of executable files and applications launched by the user (UserAssist)|What application did the user launch and when have they been used last? This gives us an insight into possible persistence targets.|

# Credits
https://github.com/louietan/LnkParser (for the logic to retrieve the target destinations of LNK files)
https://raw.githubusercontent.com/kacos2000/Jumplist-Browser/ (for the AppID list)
https://github.com/EricZimmerman/JLECmd/ (for the custom destination parsing logic)
https://github.com/EricZimmerman/RegistryPlugins/blob/master/RegistryPlugin.UserAssist/UserAssist.cs (UserAssist ROT13 Function)

# Detection

The repository contains a YARA rule to detect the unmodified pre-compiled binary.
The blog post [From Evidence to Advantage: Leveraging Incident Response Artifacts for Red Team Engagements](BLOG-LINK) contains more indicators that can be used as an idea for further detections.

# About the author

Steffen Rogge is a Cyber Security Consultant at NVISO, where he mostly conducts Red Team / Purple Team assessments with a special focus on TIBER engagements.
This enables companies to evaluate their existing defenses against emulated Advanced Persistent Threat (APT) campaigns.

If you do have a any questions regarding the tool or our services, feel free to reach out on [LinkedIn](https://www.linkedin.com/in/steffenrogge).
