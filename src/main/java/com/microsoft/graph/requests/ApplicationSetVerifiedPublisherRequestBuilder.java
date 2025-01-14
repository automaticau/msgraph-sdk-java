// Template Source: BaseMethodRequestBuilder.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests;
import com.microsoft.graph.requests.ApplicationSetVerifiedPublisherRequest;
import com.microsoft.graph.models.Application;

import com.microsoft.graph.http.BaseActionRequestBuilder;
import com.microsoft.graph.models.ApplicationSetVerifiedPublisherParameterSet;
import com.microsoft.graph.core.IBaseClient;
import com.google.gson.JsonElement;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Application Set Verified Publisher Request Builder.
 */
public class ApplicationSetVerifiedPublisherRequestBuilder extends BaseActionRequestBuilder<Application> {

    /**
     * The request builder for this ApplicationSetVerifiedPublisher
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public ApplicationSetVerifiedPublisherRequestBuilder(@Nonnull final String requestUrl, @Nonnull final IBaseClient<?> client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions);
    }
    private ApplicationSetVerifiedPublisherParameterSet body;
    /**
     * The request builder for this ApplicationSetVerifiedPublisher
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     * @param parameters     the parameters for the service method
     */
    public ApplicationSetVerifiedPublisherRequestBuilder(@Nonnull final String requestUrl, @Nonnull final IBaseClient<?> client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions, @Nonnull final ApplicationSetVerifiedPublisherParameterSet parameters) {
        super(requestUrl, client, requestOptions);
        this.body = parameters;
    }

    /**
     * Creates the ApplicationSetVerifiedPublisherRequest
     *
     * @param requestOptions the options for the request
     * @return the ApplicationSetVerifiedPublisherRequest instance
     */
    @Nonnull
    public ApplicationSetVerifiedPublisherRequest buildRequest(@Nullable final com.microsoft.graph.options.Option... requestOptions) {
        return buildRequest(getOptions(requestOptions));
    }

    /**
     * Creates the ApplicationSetVerifiedPublisherRequest with specific requestOptions instead of the existing requestOptions
     *
     * @param requestOptions the options for the request
     * @return the ApplicationSetVerifiedPublisherRequest instance
     */
    @Nonnull
    public ApplicationSetVerifiedPublisherRequest buildRequest(@Nonnull final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        final ApplicationSetVerifiedPublisherRequest request = new ApplicationSetVerifiedPublisherRequest(
                getRequestUrl(),
                getClient(),
                requestOptions);
        request.body = this.body;
        return request;
    }
}
