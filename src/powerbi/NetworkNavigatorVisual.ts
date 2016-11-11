/*
 * Copyright (c) Microsoft
 * All rights reserved.
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */
import { NetworkNavigator as NetworkNavigatorImpl } from "../NetworkNavigator";
import { INetworkNavigatorNode } from "../models";
import * as CONSTANTS from "../constants";
import { INetworkNavigatorSelectableNode, INetworkNavigatorVisualSettings } from "./models";
import { Visual, UpdateType } from "essex.powerbi.base";
import IVisualHostServices = powerbi.IVisualHostServices;
import VisualCapabilities = powerbi.VisualCapabilities;
import VisualInitOptions = powerbi.VisualInitOptions;
import VisualUpdateOptions = powerbi.VisualUpdateOptions;
import IInteractivityService = powerbi.visuals.IInteractivityService;
import InteractivityService = powerbi.visuals.InteractivityService;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import SelectionId = powerbi.visuals.SelectionId;
import utility = powerbi.visuals.utility;
import NetworkNavigatorVisualState from "./state";
import { StatefulVisual } from "pbi-stateful";
import convert from "./convert";

/* tslint:disable */
const MY_CSS_MODULE = require("!css!sass!./css/NetworkNavigatorVisual.scss");

// PBI Swallows these
const EVENTS_TO_IGNORE = "mousedown mouseup click focus blur input pointerdown pointerup touchstart touchmove touchdown";

import { DATA_ROLES } from "./constants";
import { DEFAULT_SETTINGS } from "./defaults";
import capabilities from "./capabilities";

/* tslint:enable */
declare var _: any;

@Visual(require("../build").output.PowerBI)
export default class NetworkNavigator extends StatefulVisual<NetworkNavigatorVisualState> {

    public static capabilities: VisualCapabilities = capabilities;
    private myNetworkNavigator: NetworkNavigatorImpl;
    private host: IVisualHostServices;
    private interactivityService: IInteractivityService;
    private listener: { destroy: Function; };
    private _internalState: NetworkNavigatorVisualState;

    /**
     * The selection manager
     */
    private selectionManager: utility.SelectionManager;

    private settings: INetworkNavigatorVisualSettings = $.extend(true, {}, DEFAULT_SETTINGS);

    /**
     * Gets called when a node is selected
     */
    private onNodeSelected = _.debounce((node: INetworkNavigatorSelectableNode) => {
        /* tslint:disable */
        let filter: any = null;
        /* tslint:enable */
        if (node) {
            filter = powerbi.data.SemanticFilter.fromSQExpr(node.filterExpr);
            this.selectionManager.select(node.identity, false);
        } else {
            this.selectionManager.clear();
        }

        let objects: powerbi.VisualObjectInstancesToPersist = { };
        if (filter) {
            $.extend(objects, {
                merge: [
                    <VisualObjectInstance>{
                        objectName: "general",
                        selector: undefined,
                        properties: {
                            "filter": filter
                        },
                    },
                ],
            });
        } else {
            $.extend(objects, {
                remove: [
                    <VisualObjectInstance>{
                        objectName: "general",
                        selector: undefined,
                        properties: {
                            "filter": filter
                        },
                    },
                ],
            });
        }

        this.host.persistProperties(objects);
    }, 100);

    /**
     * Constructor for the network navigator
     */
    constructor(noCss = false) {
        super("NetworkNavigator", noCss);

        const className = MY_CSS_MODULE && MY_CSS_MODULE.locals && MY_CSS_MODULE.locals.className;
        if (className) {
            this.element.addClass(className);
        }

        this._internalState = NetworkNavigatorVisualState.create();
    }

    public generateState() {
        return this._internalState.toJSONObject();
    }

    public onSetState(state: NetworkNavigatorVisualState) {
        if (state) {
            this._internalState = this._internalState.receive(state);
        }
    }

    public getCustomCssModules(): string[] {
        return [MY_CSS_MODULE];
    }

    /**
     * Gets the template for this visual
     */
    public get template() {
        return `<div id="node_graph" style= "height: 100%;"> </div>`;
    }

    /** This is called once when the visual is initialially created */
    public onInit(options: VisualInitOptions): void {
        this.myNetworkNavigator = new NetworkNavigatorImpl(this.element.find("#node_graph"), 500, 500);
        this.host = options.host;
        this.interactivityService = new InteractivityService(this.host);
        this.attachEvents();
        this.selectionManager = new utility.SelectionManager({ hostServices: this.host });
    }

    /** Update is called for data updates, resizes & formatting changes */
    public onUpdate(options: VisualUpdateOptions, type: UpdateType) {
        let dataView = options.dataViews && options.dataViews.length && options.dataViews[0];
        let dataViewTable = dataView && dataView.table;
        let forceReloadData = false;

        if (type & UpdateType.Settings) {
            forceReloadData = this.updateSettings(options);
        }
        if (type & UpdateType.Resize) {
            this.myNetworkNavigator.dimensions = { width: options.viewport.width, height: options.viewport.height };
            this.element.css({ width: options.viewport.width, height: options.viewport.height });
        }
        if (type & UpdateType.Data || forceReloadData) {
            if (dataViewTable) {
                const newData = convert(dataView, this.settings);
                this.myNetworkNavigator.setData(newData);
            } else {
                this.myNetworkNavigator.setData({
                    links: [],
                    nodes: [],
                });
            }
        }

        const data = this.myNetworkNavigator.getData();
        const nodes = data && data.nodes;
        const selectedIds = this.selectionManager.getSelectionIds();
        if (nodes && nodes.length) {
            let updated = false;
            nodes.forEach((n) => {
                let isSelected =
                    !!_.find(selectedIds, (id: SelectionId) => id.equals((<INetworkNavigatorSelectableNode>n).identity));
                if (isSelected !== n.selected) {
                    n.selected = isSelected;
                    updated = true;
                }
            });

            if (updated) {
                this.myNetworkNavigator.redrawSelection();
            }
        }

        this.myNetworkNavigator.redrawLabels();
    }

    /**
     * Enumerates the instances for the objects that appear in the power bi panel
     */
    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] {
        let instances = super.enumerateObjectInstances(options) || [{
            /* tslint:disable */
            selector: null,
            /* tslint:enable */
            objectName: options.objectName,
            properties: {},
        }, ];

        if (options.objectName === "layout") {
            const { layout } = this.settings;
            // autoClamp
            Object.keys(layout).forEach((name: string) => {
                if (CONSTANTS[name]) {
                    const { min, max } = CONSTANTS[name];
                    const value = layout[name];
                    layout[name] = Math.min(max, Math.max(min, value));
                }
            });
        }

        $.extend(true, instances[0].properties, this.settings[options.objectName]);

        if (options.objectName === "general") {
            instances[0].properties["textSize"] = this.myNetworkNavigator.configuration.fontSizePT;
        }
        return instances as VisualObjectInstance[];
    }

    /**
     * Handles updating of the settings
     */
    private updateSettings(options: VisualUpdateOptions): boolean {
        // There are some changes to the options
        let dataView = options.dataViews && options.dataViews.length && options.dataViews[0];
        if (dataView && dataView.metadata) {
            const oldSettings = $.extend(true, {}, this.settings);
            const newObjects = dataView.metadata.objects;
            const layoutObjs = newObjects && newObjects["layout"];
            const generalObjs = newObjects && newObjects["general"];

            // Merge in the settings
            $.extend(true, this.settings, DEFAULT_SETTINGS, newObjects ? newObjects : {}, {
                layout: {
                    fontSizePT: generalObjs && generalObjs["textSize"],
                    defaultLabelColor: layoutObjs && layoutObjs["defaultLabelColor"] && layoutObjs["defaultLabelColor"].solid.color,
                },
            });

            // Remove the general section, added by the above statement
            delete this.settings["general"];

            // There were some changes to the layout
            if (!_.isEqual(oldSettings, this.settings)) {
                this.myNetworkNavigator.configuration = $.extend(true, {}, this.settings.search, this.settings.layout);
            }

            if (oldSettings.layout.maxNodeCount !== this.settings.layout.maxNodeCount) {
                return true;
            }
        }
        return false;
    }

    /**
     * Returns if all the properties in the first object are present and equal to the ones in the super set
     */
    private objectIsSubset(set: Object, superSet: Object) {
        if (_.isObject(set)) {
            return _.any(_.keys(set), (key: string) => !this.objectIsSubset(set[key], superSet[key]));
        }
        return set === superSet;
    }

    /**
     * Attaches the line up events to lineup
     */
    private attachEvents() {
        if (this.myNetworkNavigator) {
            // Cleans up events
            if (this.listener) {
                this.listener.destroy();
            }
            this.listener =
                this.myNetworkNavigator.events.on("selectionChanged", (node: INetworkNavigatorNode) => this.onNodeSelected(node));

            // HAX: I am a strong, independent element and I don't need no framework tellin me how much focus I can have
            this.element.find(".filter-box input").on(EVENTS_TO_IGNORE, (e) => e.stopPropagation());
        }
    }
}
