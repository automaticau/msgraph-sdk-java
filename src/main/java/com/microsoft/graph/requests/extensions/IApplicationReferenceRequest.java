// Template Source: IBaseEntityReferenceRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.Application;
import com.microsoft.graph.models.extensions.KeyCredential;
import com.microsoft.graph.models.extensions.PasswordCredential;
import com.microsoft.graph.requests.extensions.IDirectoryObjectWithReferenceRequestBuilder;
import com.microsoft.graph.requests.extensions.IExtensionPropertyCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.IExtensionPropertyRequestBuilder;
import com.microsoft.graph.requests.extensions.IHomeRealmDiscoveryPolicyCollectionWithReferencesRequestBuilder;
import com.microsoft.graph.requests.extensions.IHomeRealmDiscoveryPolicyWithReferenceRequestBuilder;
import com.microsoft.graph.requests.extensions.IDirectoryObjectCollectionWithReferencesRequestBuilder;
import com.microsoft.graph.requests.extensions.IAppRoleAssignmentWithReferenceRequestBuilder;
import com.microsoft.graph.requests.extensions.IEndpointWithReferenceRequestBuilder;
import com.microsoft.graph.requests.extensions.IServicePrincipalWithReferenceRequestBuilder;
import com.microsoft.graph.requests.extensions.IUserWithReferenceRequestBuilder;
import com.microsoft.graph.requests.extensions.ITokenIssuancePolicyCollectionWithReferencesRequestBuilder;
import com.microsoft.graph.requests.extensions.ITokenIssuancePolicyWithReferenceRequestBuilder;
import com.microsoft.graph.requests.extensions.ITokenLifetimePolicyCollectionWithReferencesRequestBuilder;
import com.microsoft.graph.requests.extensions.ITokenLifetimePolicyWithReferenceRequestBuilder;
import java.util.Arrays;
import java.util.EnumSet;

import com.microsoft.graph.models.extensions.Application;
import com.microsoft.graph.http.IHttpRequest;
import com.microsoft.graph.serializer.IJsonBackedObject;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The interface for the Application Reference Request.
 */
public interface IApplicationReferenceRequest extends IHttpRequest {

    void delete(final ICallback<? super Application> callback);

    Application delete() throws ClientException;

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
    IApplicationReferenceRequest select(final String value);

    /**
     * Sets the expand clause for the request
     *
     * @param value the expand clause
     * @return the updated request
     */
    IApplicationReferenceRequest expand(final String value);

    /**
     * Puts the Application
     *
     * @param srcApplication the Application to PUT
     * @param callback the callback to be called after success or failure
     */
    void put(Application srcApplication, final ICallback<? super Application> callback);

    /**
     * Puts the Application
     *
     * @param srcApplication the Application to PUT
     * @return the Application
     * @throws ClientException an exception occurs if there was an error while the request was sent
     */
    Application put(Application srcApplication) throws ClientException;
}
