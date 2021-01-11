// Template Source: BaseMethodRequestBuilder.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;
import com.microsoft.graph.requests.extensions.EventDeclineRequest;
import com.microsoft.graph.models.extensions.Event;

import com.microsoft.graph.http.BaseActionRequestBuilder;
import com.microsoft.graph.models.extensions.EventDeclineParameterSet;
import com.microsoft.graph.core.IBaseClient;
import com.google.gson.JsonElement;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Event Decline Request Builder.
 */
public class EventDeclineRequestBuilder extends BaseActionRequestBuilder<Event> {

    private EventDeclineParameterSet body;
    /**
     * The request builder for this EventDecline
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     * @param parameters     the parameters for the service method
     */
    public EventDeclineRequestBuilder(@Nonnull final String requestUrl, @Nonnull final IBaseClient client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions, @Nonnull final EventDeclineParameterSet parameters) {
        super(requestUrl, client, requestOptions);
        this.body = parameters;
    }

    /**
     * Creates the EventDeclineRequest
     *
     * @param requestOptions the options for the request
     * @return the EventDeclineRequest instance
     */
    @Nonnull
    public EventDeclineRequest buildRequest(@Nullable final com.microsoft.graph.options.Option... requestOptions) {
        return buildRequest(getOptions(requestOptions));
    }

    /**
     * Creates the EventDeclineRequest with specific requestOptions instead of the existing requestOptions
     *
     * @param requestOptions the options for the request
     * @return the EventDeclineRequest instance
     */
    @Nonnull
    public EventDeclineRequest buildRequest(@Nonnull final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        EventDeclineRequest request = new EventDeclineRequest(
                getRequestUrl(),
                getClient(),
                requestOptions,
                this.body
        );

            return request;
    }
}
