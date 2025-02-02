// Template Source: BaseMethodRequestBuilder.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests;
import com.microsoft.graph.requests.DeviceCompliancePolicyScheduleActionsForRulesRequest;
import com.microsoft.graph.models.DeviceCompliancePolicy;
import com.microsoft.graph.models.DeviceComplianceScheduledActionForRule;
import com.microsoft.graph.http.BaseActionRequestBuilder;
import com.microsoft.graph.models.DeviceCompliancePolicyScheduleActionsForRulesParameterSet;
import com.microsoft.graph.core.IBaseClient;
import com.google.gson.JsonElement;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Device Compliance Policy Schedule Actions For Rules Request Builder.
 */
public class DeviceCompliancePolicyScheduleActionsForRulesRequestBuilder extends BaseActionRequestBuilder<DeviceCompliancePolicy> {

    /**
     * The request builder for this DeviceCompliancePolicyScheduleActionsForRules
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public DeviceCompliancePolicyScheduleActionsForRulesRequestBuilder(@Nonnull final String requestUrl, @Nonnull final IBaseClient<?> client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions);
    }
    private DeviceCompliancePolicyScheduleActionsForRulesParameterSet body;
    /**
     * The request builder for this DeviceCompliancePolicyScheduleActionsForRules
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     * @param parameters     the parameters for the service method
     */
    public DeviceCompliancePolicyScheduleActionsForRulesRequestBuilder(@Nonnull final String requestUrl, @Nonnull final IBaseClient<?> client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions, @Nonnull final DeviceCompliancePolicyScheduleActionsForRulesParameterSet parameters) {
        super(requestUrl, client, requestOptions);
        this.body = parameters;
    }

    /**
     * Creates the DeviceCompliancePolicyScheduleActionsForRulesRequest
     *
     * @param requestOptions the options for the request
     * @return the DeviceCompliancePolicyScheduleActionsForRulesRequest instance
     */
    @Nonnull
    public DeviceCompliancePolicyScheduleActionsForRulesRequest buildRequest(@Nullable final com.microsoft.graph.options.Option... requestOptions) {
        return buildRequest(getOptions(requestOptions));
    }

    /**
     * Creates the DeviceCompliancePolicyScheduleActionsForRulesRequest with specific requestOptions instead of the existing requestOptions
     *
     * @param requestOptions the options for the request
     * @return the DeviceCompliancePolicyScheduleActionsForRulesRequest instance
     */
    @Nonnull
    public DeviceCompliancePolicyScheduleActionsForRulesRequest buildRequest(@Nonnull final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        final DeviceCompliancePolicyScheduleActionsForRulesRequest request = new DeviceCompliancePolicyScheduleActionsForRulesRequest(
                getRequestUrl(),
                getClient(),
                requestOptions);
        request.body = this.body;
        return request;
    }
}
