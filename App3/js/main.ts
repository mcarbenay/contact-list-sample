﻿/// <reference path="../scripts/typings/jquery/jquery.d.ts" />

module ListeContacts {

    /**
     * Représente un contact récupéré depuis
     * votre back-end
     */
    class ContactBackend {
        public remoteId: string;
        public name: string;
        public email: string;
    }

    export class ListeContactSample {

        private static listName: string = "exemple de liste";

        public static MettreAJour() {
            // on obtient une ref. sur le store en mode "écriture"
            Windows.ApplicationModel.Contacts.ContactManager.requestStoreAsync(
                Windows.ApplicationModel.Contacts.ContactStoreAccessType.appContactsReadWrite
            ).done(function (store) {
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
        }

        public static Supprimer() {
            // on obtient une ref. sur le store en mode "écriture"
            Windows.ApplicationModel.Contacts.ContactManager.requestStoreAsync(
                Windows.ApplicationModel.Contacts.ContactStoreAccessType.appContactsReadWrite
            ).done(function (store) {
                // on énumère les listes pour vérifier si la notre existe
                // déjà : si oui, on l'update sinon on la créé
                store.findContactListsAsync().done(function (lists) {
                    var laListe : Windows.ApplicationModel.Contacts.ContactList = null;
                    for (var i = 0; i < lists.length; i++) {
                        if (lists[i].displayName == ListeContactSample.listName) {
                            laListe = lists[i];
                        }
                    }
                    if (laListe != null) {
                        laListe.deleteAsync().done(function () {

                        });
                    }
                });
            });
        }


        private static createContactList(store : Windows.ApplicationModel.Contacts.ContactStore) {
            store.createContactListAsync(ListeContactSample.listName).done(function (laListe) {
                laListe.otherAppReadAccess = Windows.ApplicationModel.Contacts.ContactListOtherAppReadAccess.full;
                laListe.otherAppWriteAccess = Windows.ApplicationModel.Contacts.ContactListOtherAppWriteAccess.none;
                laListe.saveAsync().done(function () {
                    ListeContactSample.refreshContactList(laListe);
                });
            });
        }


        /**
         * Obtient les données du back-end
         * pour l'exemple 'obtient les données en dur' :)
         * @param success le callback à appeler après récupération des données
         */
        private static getAllFromBackEnd(success: (ctcs: ContactBackend[]) => void) {
            var ret = new Array<ContactBackend>();

            ret.push({
                remoteId: "7905037E-794B-4FA2-8411-1B85E58D6AD0",
                name: "Contact 1",
                email : "contact1@mondomaine.com"
            });

            ret.push({
                remoteId: "85E0D6AF-2ACD-461E-9516-E6DF7C76384B",
                name: "Contact 2",
                email: "contact2@monautredomaine.com"
            });

            //ret.push({
            //    remoteId: "E862CB70-6250-40D7-956D-300D1E6BA134",
            //    name: "Contact 3",
            //    email: "contact3@encoreunautredomain.com"
            //});

            if (success != null) {
                success(ret);
            }

        }

        private static refreshContactList(laListe: Windows.ApplicationModel.Contacts.ContactList) {
            ListeContactSample.getAllFromBackEnd(function (ctcs) {
                for (var i = 0; i < ctcs.length; i++) {
                    ListeContactSample.refreshContact(laListe, ctcs[i]);
                }

                // premier jet de la suppression
                laListe.getContactReader().readBatchAsync().done(function (batch) {
                    for (var i = 0; i < batch.contacts.length; i++) {
                        var found = false;
                        for (var j = 0; j < ctcs.length; j++) {
                            if (ctcs[j].remoteId == batch.contacts[i].remoteId) {
                                found = true;
                                break;
                            }
                        }
                        if (!found) {
                            laListe.deleteContactAsync(batch.contacts[i]);
                        }
                    }
                });
            });
        }
        private static refreshContact(laListe: Windows.ApplicationModel.Contacts.ContactList,
            leContact: ContactBackend) {
            // obtient le contact depuis son id "distant"
            laListe.getContactFromRemoteIdAsync(leContact.remoteId).done(function (ctc) {
                // si il n'existe pas, il faut le créer
                if (ctc == null) {
                    ctc = new Windows.ApplicationModel.Contacts.Contact();
                    ctc.remoteId = leContact.remoteId;
                } else {
                    ctc.emails.clear();
                }
                // puis mettre à jour ses données
                ctc.name = leContact.name;
                if (leContact.email != null && leContact.email != "") {
                    var email = new Windows.ApplicationModel.Contacts.ContactEmail();
                    email.address = leContact.email;
                    email.kind = Windows.ApplicationModel.Contacts.ContactEmailKind.work;
                    ctc.emails.push(email);
                }
                // et finalement l'enregistrer
                laListe.saveContactAsync(ctc).done(function () {
                    
                });
            })
        }


    }
}

$(document).ready(() => {
    $("#btnMakeList").click(() => ListeContacts.ListeContactSample.MettreAJour());
    $("#btnDeleteList").click(() => ListeContacts.ListeContactSample.Supprimer());
});
