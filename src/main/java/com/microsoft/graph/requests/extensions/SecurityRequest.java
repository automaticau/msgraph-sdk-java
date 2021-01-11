// Template Source: BaseEntityRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.Security;
import com.microsoft.graph.requests.extensions.AlertCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.AlertRequestBuilder;
import com.microsoft.graph.requests.extensions.SecureScoreControlProfileCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.SecureScoreControlProfileRequestBuilder;
import com.microsoft.graph.requests.extensions.SecureScoreCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.SecureScoreRequestBuilder;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequest;
import com.microsoft.graph.http.HttpMethod;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Security Request.
 */
public class SecurityRequest extends BaseRequest<Security> {
	
    /**
     * The request for the Security
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public SecurityRequest(@Nonnull final String requestUrl, @Nonnull final IBaseClient client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, Security.class);
    }

    /**
     * Gets the Security from the service
     *
     * @param callback the callback to be called after success or failure
     */
    public void get(@Nonnull final ICallback<? super Security> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the Security from the service
     *
     * @return the Security from the request
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public Security get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service
     *
     * @param callback the callback when the deletion action has completed
     */
    public void delete(@Nonnull final ICallback<? super Security> callback) {
        send(HttpMethod.DELETE, callback, null);
    }

    /**
     * Delete this item from the service
     *
     * @throws ClientException if there was an exception during the delete operation
     */
    public void delete() throws ClientException {
        send(HttpMethod.DELETE, null);
    }

    /**
     * Patches this Security with a source
     *
     * @param sourceSecurity the source object with updates
     * @param callback the callback to be called after success or failure
     */
    public void patch(@Nonnull final Security sourceSecurity, @Nonnull final ICallback<? super Security> callback) {
        send(HttpMethod.PATCH, callback, sourceSecurity);
    }

    /**
     * Patches this Security with a source
     *
     * @param sourceSecurity the source object with updates
     * @return the updated Security
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public Security patch(@Nonnull final Security sourceSecurity) throws ClientException {
        return send(HttpMethod.PATCH, sourceSecurity);
    }

    /**
     * Creates a Security with a new object
     *
     * @param newSecurity the new object to create
     * @param callback the callback to be called after success or failure
     */
    public void post(@Nonnull final Security newSecurity, @Nonnull final ICallback<? super Security> callback) {
        send(HttpMethod.POST, callback, newSecurity);
    }

    /**
     * Creates a Security with a new object
     *
     * @param newSecurity the new object to create
     * @return the created Security
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public Security post(@Nonnull final Security newSecurity) throws ClientException {
        return send(HttpMethod.POST, newSecurity);
    }

    /**
     * Creates a Security with a new object
     *
     * @param newSecurity the object to create/update
     * @param callback the callback to be called after success or failure
     */
    public void put(@Nonnull final Security newSecurity, @Nonnull final ICallback<? super Security> callback) {
        send(HttpMethod.PUT, callback, newSecurity);
    }

    /**
     * Creates a Security with a new object
     *
     * @param newSecurity the object to create/update
     * @return the created Security
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public Security put(@Nonnull final Security newSecurity) throws ClientException {
        return send(HttpMethod.PUT, newSecurity);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
     @Nonnull
     public SecurityRequest select(@Nonnull final String value) {
         addSelectOption(value);
         return this;
     }

    /**
     * Sets the expand clause for the request
     *
     * @param value the expand clause
     * @return the updated request
     */
     @Nonnull
     public SecurityRequest expand(@Nonnull final String value) {
         addExpandOption(value);
         return this;
     }

}

