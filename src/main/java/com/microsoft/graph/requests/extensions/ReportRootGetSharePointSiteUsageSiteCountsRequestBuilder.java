// Template Source: BaseMethodRequestBuilder.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;
import com.microsoft.graph.requests.extensions.ReportRootGetSharePointSiteUsageSiteCountsRequest;
import com.microsoft.graph.models.extensions.ReportRoot;
import com.microsoft.graph.models.extensions.Report;
import com.microsoft.graph.http.BaseFunctionRequestBuilder;
import com.microsoft.graph.models.extensions.ReportRootGetSharePointSiteUsageSiteCountsParameterSet;
import com.microsoft.graph.core.IBaseClient;
import com.google.gson.JsonElement;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Report Root Get Share Point Site Usage Site Counts Request Builder.
 */
public class ReportRootGetSharePointSiteUsageSiteCountsRequestBuilder extends BaseFunctionRequestBuilder<Report> {

    /**
     * The request builder for this ReportRootGetSharePointSiteUsageSiteCounts
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     * @param parameters     the parameters for the service method
     */
    public ReportRootGetSharePointSiteUsageSiteCountsRequestBuilder(@Nonnull final String requestUrl, @Nonnull final IBaseClient client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions, @Nonnull final ReportRootGetSharePointSiteUsageSiteCountsParameterSet parameters) {
        super(requestUrl, client, requestOptions);
        if(parameters != null) {
            functionOptions = parameters.getFunctionOptions();
        }
    }

    /**
     * Creates the ReportRootGetSharePointSiteUsageSiteCountsRequest
     *
     * @param requestOptions the options for the request
     * @return the ReportRootGetSharePointSiteUsageSiteCountsRequest instance
     */
    @Nonnull
    public ReportRootGetSharePointSiteUsageSiteCountsRequest buildRequest(@Nullable final com.microsoft.graph.options.Option... requestOptions) {
        return buildRequest(getOptions(requestOptions));
    }

    /**
     * Creates the ReportRootGetSharePointSiteUsageSiteCountsRequest with specific requestOptions instead of the existing requestOptions
     *
     * @param requestOptions the options for the request
     * @return the ReportRootGetSharePointSiteUsageSiteCountsRequest instance
     */
    @Nonnull
    public ReportRootGetSharePointSiteUsageSiteCountsRequest buildRequest(@Nonnull final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        ReportRootGetSharePointSiteUsageSiteCountsRequest request = new ReportRootGetSharePointSiteUsageSiteCountsRequest(
                getRequestUrl(),
                getClient(),
                requestOptions
        );

            for (com.microsoft.graph.options.FunctionOption option : functionOptions) {
            request.addFunctionOption(option);
        }
        return request;
    }
}
