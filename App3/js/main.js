﻿/// <reference path="../scripts/typings/jquery/jquery.d.ts" />
var ListeContacts;
(function (ListeContacts) {
    /**
     * Représente un contact récupéré depuis
     * votre back-end
     */
    var ContactBackend = (function () {
        function ContactBackend() {
        }
        return ContactBackend;
    }());
    var ListeContactSample = (function () {
        function ListeContactSample() {
        }
        ListeContactSample.MettreAJour = function () {
            // on obtient une ref. sur le store en mode "écriture"
            Windows.ApplicationModel.Contacts.ContactManager.requestStoreAsync(Windows.ApplicationModel.Contacts.ContactStoreAccessType.appContactsReadWrite).done(function (store) {
                // on énumère les listes pour vérifier si la notre existe
                // déjà : si oui, on l'update sinon on la créé
                store.findContactListsAsync().done(function (lists) {
                    var laListe = null;
                    for (var i = 0; i < lists.length; i++) {
                        if (lists[i].displayName == ListeContactSample.listName) {
                            laListe = lists[i];
                        }
                    }
                    if (laListe == null) {
                        ListeContactSample.createContactList(store);
                    }
                    else {
                        ListeContactSample.refreshContactList(laListe);
                    }
                });
            });
        };
        ListeContactSample.createContactList = function (store) {
            store.createContactListAsync(ListeContactSample.listName).done(function (laListe) {
                laListe.otherAppReadAccess = Windows.ApplicationModel.Contacts.ContactListOtherAppReadAccess.full;
                laListe.otherAppWriteAccess = Windows.ApplicationModel.Contacts.ContactListOtherAppWriteAccess.none;
                laListe.saveAsync().done(function () {
                    ListeContactSample.refreshContactList(laListe);
                });
            });
        };
        ListeContactSample.refreshContactList = function (laListe) {
            ListeContactSample.getAllFromBackEnd(function (ctcs) {
                for (var i = 0; i < ctcs.length; i++) {
                    ListeContactSample.refreshContact(laListe, ctcs[i]);
                }
            });
        };
        /**
         * Obtient les données du back-end
         * pour l'exemple 'obtient les données en dur' :)
         * @param success le callback à appeler après récupération des données
         */
        ListeContactSample.getAllFromBackEnd = function (success) {
            var ret = new Array();
            ret.push({
                remoteId: "7905037E-794B-4FA2-8411-1B85E58D6AD0",
                name: "Contact 1",
                email: "contact1@mondomaine.com"
            });
            ret.push({
                remoteId: "85E0D6AF-2ACD-461E-9516-E6DF7C76384B",
                name: "Contact 2",
                email: "contact2@monautredomaine.com"
            });
            if (success != null) {
                success(ret);
            }
        };
        ListeContactSample.refreshContact = function (laListe, leContact) {
            laListe.getContactFromRemoteIdAsync(leContact.remoteId).done(function (ctc) {
                if (ctc == null) {
                    ctc = new Windows.ApplicationModel.Contacts.Contact();
                    ctc.remoteId = leContact.remoteId;
                }
                else {
                    ctc.emails.clear();
                }
                ctc.name = leContact.name;
                if (leContact.email != null && leContact.email != "") {
                    var email = new Windows.ApplicationModel.Contacts.ContactEmail();
                    email.address = leContact.email;
                    email.kind = Windows.ApplicationModel.Contacts.ContactEmailKind.work;
                    ctc.emails.push(email);
                }
                laListe.saveContactAsync(ctc).done(function () {
                });
            });
        };
        return ListeContactSample;
    }());
    ListeContactSample.listName = "exemple de liste";
    ListeContacts.ListeContactSample = ListeContactSample;
})(ListeContacts || (ListeContacts = {}));
$(document).ready(function () {
    $("#btnMakeList").click(function () { return ListeContacts.ListeContactSample.MettreAJour(); });
});
//# sourceMappingURL=c:/temp/App3/App3/js/main.js.map