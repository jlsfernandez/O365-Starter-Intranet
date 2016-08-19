// ========================================
// Contextual Menu Component View Model
// ========================================

import { NavigationViewModel } from "../shared/navigation.viewmodel";
import { UtilityModule } from "../core/utility";
import { NavigationNode } from "../shared/navigationnode";
import * as pnp from "sp-pnp-js";
import "pubsub-js";

export class ContextualMenuViewModel extends NavigationViewModel {

    public siteMapFieldName: string;
    public utilityModule: UtilityModule;
    public wait: KnockoutObservable<boolean>;

    constructor(params: any) {

        super();

        this.utilityModule = new UtilityModule();

        // The internal name for the site map taxonomy field
        this.siteMapFieldName = params.siteMapFieldName;

        this.wait = ko.observable(true);

        // Collapse events
        $("#contextualmenu").on("hide.bs.collapse", (event) => {

            let iconElt = $("[data-target='#" + event.target.id + "']").find("i");
            if (iconElt) {
                iconElt.removeClass("ms-Icon--chevronUp");
                iconElt.addClass("ms-Icon--chevronDown");
            }
        });

        $("#contextualmenu").on("show.bs.collapse", (event) => {

            event.stopPropagation();

            // Get the parent with the data-target attribute equals to my id.
            let iconElt = $("[data-target='#" + event.target.id + "']").find("i");
            if (iconElt) {
                iconElt.removeClass("ms-Icon--chevronDown");
                iconElt.addClass("ms-Icon--chevronUp");
            }
        });

        // Subscribe to the main menu nodes
        PubSub.subscribe("navigationNodes", (msg, data) => {

            let navigationTree = data.nodes;

            pnp.sp.web.lists.getByTitle("Pages").items.getById(_spPageContextInfo.pageItemId).select(this.siteMapFieldName).get().then((item) => {

                    let siteMapTermGuid = item[this.siteMapFieldName];
                    let currentNode: NavigationNode = undefined;

                    if (siteMapTermGuid) {

                        // 1: Search for this guid in the site map
                        currentNode = this.utilityModule.getNodeByTermId(navigationTree, siteMapTermGuid.TermGuid);
                    }

                    if (currentNode === undefined) {

                        // 2: Get the navigation node according to the current URL   
                        currentNode = this.utilityModule.getNodeByUrl(navigationTree, window.location.pathname);
                    }

                    if (currentNode !== undefined) {

                        // If there is no 'ParentId', this is a root term
                        if (currentNode.ParentId !== null) {

                            navigationTree = this.utilityModule.getNodeByTermId(navigationTree, new SP.Guid(currentNode.ParentId));

                            if (navigationTree.ChildNodes.length > 0) {

                                // Display all siblings and child nodes from the current node (just like the CSOM results)
                                // Siblings = children of my own parent ;)
                                navigationTree = navigationTree.ChildNodes;
                            }
                        }

                    } else {

                        pnp.log.write("[Contextual Menu] Unable to determine the current position in the site map", pnp.log.LogLevel.Warning);
                    }

                    this.initialize(navigationTree);
                    this.wait(false);

                    if (currentNode !== undefined) {

                        this.setCurrentNode(new SP.Guid(currentNode.Id));
                    }

            }).catch((errorMesssage) => {

                pnp.log.write(errorMesssage, pnp.log.LogLevel.Error);
            });
        });
    }
}
