// Template Source: BaseEntityRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.Presence;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequest;
import com.microsoft.graph.http.HttpMethod;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Presence Request.
 */
public class PresenceRequest extends BaseRequest<Presence> {
	
    /**
     * The request for the Presence
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public PresenceRequest(@Nonnull final String requestUrl, @Nonnull final IBaseClient client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, Presence.class);
    }

    /**
     * Gets the Presence from the service
     *
     * @param callback the callback to be called after success or failure
     */
    public void get(@Nonnull final ICallback<? super Presence> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the Presence from the service
     *
     * @return the Presence from the request
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public Presence get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service
     *
     * @param callback the callback when the deletion action has completed
     */
    public void delete(@Nonnull final ICallback<? super Presence> callback) {
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
     * Patches this Presence with a source
     *
     * @param sourcePresence the source object with updates
     * @param callback the callback to be called after success or failure
     */
    public void patch(@Nonnull final Presence sourcePresence, @Nonnull final ICallback<? super Presence> callback) {
        send(HttpMethod.PATCH, callback, sourcePresence);
    }

    /**
     * Patches this Presence with a source
     *
     * @param sourcePresence the source object with updates
     * @return the updated Presence
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public Presence patch(@Nonnull final Presence sourcePresence) throws ClientException {
        return send(HttpMethod.PATCH, sourcePresence);
    }

    /**
     * Creates a Presence with a new object
     *
     * @param newPresence the new object to create
     * @param callback the callback to be called after success or failure
     */
    public void post(@Nonnull final Presence newPresence, @Nonnull final ICallback<? super Presence> callback) {
        send(HttpMethod.POST, callback, newPresence);
    }

    /**
     * Creates a Presence with a new object
     *
     * @param newPresence the new object to create
     * @return the created Presence
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public Presence post(@Nonnull final Presence newPresence) throws ClientException {
        return send(HttpMethod.POST, newPresence);
    }

    /**
     * Creates a Presence with a new object
     *
     * @param newPresence the object to create/update
     * @param callback the callback to be called after success or failure
     */
    public void put(@Nonnull final Presence newPresence, @Nonnull final ICallback<? super Presence> callback) {
        send(HttpMethod.PUT, callback, newPresence);
    }

    /**
     * Creates a Presence with a new object
     *
     * @param newPresence the object to create/update
     * @return the created Presence
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public Presence put(@Nonnull final Presence newPresence) throws ClientException {
        return send(HttpMethod.PUT, newPresence);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
     @Nonnull
     public PresenceRequest select(@Nonnull final String value) {
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
     public PresenceRequest expand(@Nonnull final String value) {
         addExpandOption(value);
         return this;
     }

}

