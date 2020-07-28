
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.concurrency.*;
import com.microsoft.graph.core.*;
import com.microsoft.graph.models.extensions.*;
import com.microsoft.graph.models.generated.*;
import com.microsoft.graph.http.*;
import com.microsoft.graph.requests.extensions.*;
import com.microsoft.graph.options.*;
import com.microsoft.graph.serializer.*;

import java.util.Arrays;
import java.util.EnumSet;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Home Realm Discovery Policy Reference Request Builder.
 */
public class HomeRealmDiscoveryPolicyReferenceRequestBuilder extends BaseRequestBuilder implements IHomeRealmDiscoveryPolicyReferenceRequestBuilder {

    /**
     * The request builder for the HomeRealmDiscoveryPolicy
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public HomeRealmDiscoveryPolicyReferenceRequestBuilder(final String requestUrl, final IBaseClient client, final java.util.List<? extends Option> requestOptions) {
        super(requestUrl, client, requestOptions);
    }

    /**
     * Creates the request
     *
     * @return The IHomeRealmDiscoveryPolicyReferenceRequest instance
     */
    public IHomeRealmDiscoveryPolicyReferenceRequest buildRequest() {
        return buildRequest(getOptions());
    }

    /**
     * Creates the request with specific requestOptions instead of the existing requestOptions
     *
     * @param requestOptions the options for this request
     * @return the IHomeRealmDiscoveryPolicyReferenceRequest instance
     */
    public IHomeRealmDiscoveryPolicyReferenceRequest buildRequest(final java.util.List<? extends Option> requestOptions) {
        return new HomeRealmDiscoveryPolicyReferenceRequest(getRequestUrl(), getClient(), requestOptions);
    }
}
