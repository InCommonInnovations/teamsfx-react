import { __assign, __awaiter, __extends, __generator } from "tslib";
import React, { Component } from "react";
/**
 * Returns the CSS class name for the dashboard.
 * @returns The CSS class name for the dashboard.
 * @internal
 */
function dashboardStyle(isMobile) {
    return (__assign({ display: "grid", gap: "20px", padding: "20px", gridTemplateRows: "1fr", gridTemplateColumns: "4fr 6fr" }, (isMobile === true ? { gridTemplateColumns: "1fr", gridTemplateRows: "1fr" } : {})));
}
/**
 * The base component that provides basic functionality to create a dashboard.
 * @typeParam P The type of props.
 * @typeParam S The type of state.
 */
var BaseDashboard = /** @class */ (function (_super) {
    __extends(BaseDashboard, _super);
    /**
     * Constructor of BaseDashboard.
     * @param {Readonly<P>} props The properties for the dashboard.
     */
    function BaseDashboard(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            isMobile: undefined,
            showLogin: undefined,
            observer: undefined,
        };
        _this.ref = React.createRef();
        return _this;
    }
    /**
     * Called after the component is mounted. You can do initialization that requires DOM nodes here. You can also make network requests here if you need to load data from a remote endpoint.
     */
    BaseDashboard.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            var observer;
            var _this = this;
            return __generator(this, function (_a) {
                observer = new ResizeObserver(function (entries) {
                    for (var _i = 0, entries_1 = entries; _i < entries_1.length; _i++) {
                        var entry = entries_1[_i];
                        if (entry.target === _this.ref.current) {
                            var width = entry.contentRect.width;
                            _this.setState({ isMobile: width < 600 });
                        }
                    }
                });
                observer.observe(this.ref.current);
                return [2 /*return*/];
            });
        });
    };
    /**
     * Called before the component is unmounted and destroyed. You can do necessary cleanup here, such as invalidating timers, canceling network requests, or removing any DOM elements.
     */
    BaseDashboard.prototype.componentWillUnmount = function () {
        // Unobserve the dashboard div for resize events
        if (this.state.observer && this.ref.current) {
            this.state.observer.unobserve(this.ref.current);
        }
    };
    /**
     * Defines the default layout for the dashboard.
     */
    BaseDashboard.prototype.render = function () {
        return (React.createElement("div", { ref: this.ref, className: (dashboardStyle(this.state.isMobile), this.styling()) }, this.layout()));
    };
    /**
     * Override this method to define the layout of the widget in the dashboard.
     * @returns The layout of the widget in the dashboard.
     * @public
     */
    BaseDashboard.prototype.layout = function () {
        return undefined;
    };
    /**
     * Override this method to customize the dashboard style.
     * @returns The className for customizing the dashboard style.
     * @public
     */
    BaseDashboard.prototype.styling = function () {
        return null;
    };
    return BaseDashboard;
}(Component));
export { BaseDashboard };
