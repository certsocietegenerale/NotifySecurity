# What is NotifySecurity?

NotifySecurity is an Outlook add-in used to help your users to report suspicious e-mails to security teams.

It is also built as a companion tool to [Swordphish](https://github.com/certsocietegenerale/swordphish-awareness/): it will automatically detect an e-mail issued by that platform and update the "reported statistics" accordingly.

So this tool has two benefits:
* it helps your users reporting an e-mail to the right contacts with relevant information (e.g. full SMTP headers),
* it helps you tracking your users' behavior when you run a [Swordphish](https://github.com/certsocietegenerale/swordphish-awareness/) awareness campaign.

We believe that the click rate is not really a useful metric when doing awareness campaigns. Reaching 0 click is utopian, but we need to ensure that at least one target will notify a malicious e-mail to security teams.

When a victim reports an e-mail, they allow the security teams to initiate identification, containment and remediation actions; for example pivot on IOCs to identify all the targets and identify those who were compromised.

By deploying this add-in to your workstations, you make reporting suspicious e-mails a one-click process and naturally improve reporting rates.

This add-in is a very simple version, feel free to improve it and adapt it to your organization.

Once installed it will add a button in the Outlook ribbon:

![screenshot](https://github.com/certsocietegenerale/NotifySecurity/blob/master/screenshots/ribbon.png?raw=true)

And one in the contextual menu:

![screenshot](https://github.com/certsocietegenerale/NotifySecurity/blob/master/screenshots/contextual_menu.png?raw=true)

If you select an e-mail and click on this button, a notification will automatically be generated:

![screenshot](https://github.com/certsocietegenerale/NotifySecurity/blob/master/screenshots/generated_mail.png?raw=true)

# Building the project

This project can be compiled with Visual Studio Community Edition with the Office tools enabled.

You simply have to edit the project settings to adapt them to your organization and your Sworphish instance, if available.

![screenshot](https://github.com/certsocietegenerale/NotifySecurity/blob/master/screenshots/settings_place.png?raw=true)

![screenshot](https://github.com/certsocietegenerale/NotifySecurity/blob/master/screenshots/settings.png?raw=true)

# Special thanks

We would like to thank Nicolas Chaussard from ALD Automotive Security Teams for providing us the code base for this project and allowing us to release it here!