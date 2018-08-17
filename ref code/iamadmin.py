import admin
if not admin.isUserAdmin():
        admin.runAsAdmin()