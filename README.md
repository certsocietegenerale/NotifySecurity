# What is NotifySecurity ?

NotifySecurity is an Outlook companion add-in used to help your users to report suspicious mails to security teams.

It's also built in a way that it can detect automatically a mail coming from [Swordphish](https://github.com/certsocietegenerale/swordphish-awareness/) and update "reported statistics" accordingly.

So this tool has two benefits:

* it helps your users to signal a mail to the right contacts with the required informations (full headers)
* it helps you to track your users's behavior when you do an awareness campaign with [Swordphish](https://github.com/certsocietegenerale/swordphish-awareness/)

We believe that the click rate is not a really useful metric when doing awareness campaign. Reaching 0 click is utopian, but we need to ensure that at least one target will signal a malicious mail to the security teams.

When a victim reports a mail, she allows the security teams to pivot on IOCs to identify all the targets and check that no one fell into the trap.

By deploying this kind of add-in on your workstations, you'll be able to improve reporting rates.

This add-in is a very simple version, feel free to improve it and adapt it to your organization.

Once installed it will add a button in Outlook ribbon:

![screenshot](https://github.com/certsocietegenerale/NotifySecurity/blob/master/screenshots/ribbon.png?raw=true)

And one in the contextual menu:

![screenshot](https://github.com/certsocietegenerale/NotifySecurity/blob/master/screenshots/contextual_menu.png?raw=true)

If you select a mail and click on this button, a mail will automatically be generated:

![screenshot](https://github.com/certsocietegenerale/NotifySecurity/blob/master/screenshots/generated_mail.png?raw=true)

# Compilation

This project can be compiled with Visual Studio Community with the office tools.

You just have to change the project's settings to adjust it to your organization and your Sworphish instance configuration.


![screenshot](https://github.com/certsocietegenerale/NotifySecurity/blob/master/screenshots/settings_place.png?raw=true)


![screenshot](https://github.com/certsocietegenerale/NotifySecurity/blob/master/screenshots/settings.png?raw=true)


# Special Thanks

We would like to thank Nicolas Chaussard from ALD Automotive Security Teams for providing us the code base for this project and allowing us to publish it here !