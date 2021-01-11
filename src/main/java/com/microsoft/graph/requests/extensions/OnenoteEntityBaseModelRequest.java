// Template Source: BaseEntityRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.OnenoteEntityBaseModel;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequest;
import com.microsoft.graph.http.HttpMethod;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Onenote Entity Base Model Request.
 */
public class OnenoteEntityBaseModelRequest extends BaseRequest<OnenoteEntityBaseModel> {
	
    /**
     * The request for the OnenoteEntityBaseModel
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     * @param responseClass  the class of the response
     */
    public OnenoteEntityBaseModelRequest(@Nonnull final String requestUrl,
            @Nonnull final IBaseClient client,
            @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions,
            @Nonnull final Class<? extends OnenoteEntityBaseModel> responseClass) {
        super(requestUrl, client, requestOptions, responseClass);
    }

    /**
     * The request for the OnenoteEntityBaseModel
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public OnenoteEntityBaseModelRequest(@Nonnull final String requestUrl, @Nonnull final IBaseClient client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, OnenoteEntityBaseModel.class);
    }

    /**
     * Gets the OnenoteEntityBaseModel from the service
     *
     * @param callback the callback to be called after success or failure
     */
    public void get(@Nonnull final ICallback<? super OnenoteEntityBaseModel> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the OnenoteEntityBaseModel from the service
     *
     * @return the OnenoteEntityBaseModel from the request
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public OnenoteEntityBaseModel get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service
     *
     * @param callback the callback when the deletion action has completed
     */
    public void delete(@Nonnull final ICallback<? super OnenoteEntityBaseModel> callback) {
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
     * Patches this OnenoteEntityBaseModel with a source
     *
     * @param sourceOnenoteEntityBaseModel the source object with updates
     * @param callback the callback to be called after success or failure
     */
    public void patch(@Nonnull final OnenoteEntityBaseModel sourceOnenoteEntityBaseModel, @Nonnull final ICallback<? super OnenoteEntityBaseModel> callback) {
        send(HttpMethod.PATCH, callback, sourceOnenoteEntityBaseModel);
    }

    /**
     * Patches this OnenoteEntityBaseModel with a source
     *
     * @param sourceOnenoteEntityBaseModel the source object with updates
     * @return the updated OnenoteEntityBaseModel
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public OnenoteEntityBaseModel patch(@Nonnull final OnenoteEntityBaseModel sourceOnenoteEntityBaseModel) throws ClientException {
        return send(HttpMethod.PATCH, sourceOnenoteEntityBaseModel);
    }

    /**
     * Creates a OnenoteEntityBaseModel with a new object
     *
     * @param newOnenoteEntityBaseModel the new object to create
     * @param callback the callback to be called after success or failure
     */
    public void post(@Nonnull final OnenoteEntityBaseModel newOnenoteEntityBaseModel, @Nonnull final ICallback<? super OnenoteEntityBaseModel> callback) {
        send(HttpMethod.POST, callback, newOnenoteEntityBaseModel);
    }

    /**
     * Creates a OnenoteEntityBaseModel with a new object
     *
     * @param newOnenoteEntityBaseModel the new object to create
     * @return the created OnenoteEntityBaseModel
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public OnenoteEntityBaseModel post(@Nonnull final OnenoteEntityBaseModel newOnenoteEntityBaseModel) throws ClientException {
        return send(HttpMethod.POST, newOnenoteEntityBaseModel);
    }

    /**
     * Creates a OnenoteEntityBaseModel with a new object
     *
     * @param newOnenoteEntityBaseModel the object to create/update
     * @param callback the callback to be called after success or failure
     */
    public void put(@Nonnull final OnenoteEntityBaseModel newOnenoteEntityBaseModel, @Nonnull final ICallback<? super OnenoteEntityBaseModel> callback) {
        send(HttpMethod.PUT, callback, newOnenoteEntityBaseModel);
    }

    /**
     * Creates a OnenoteEntityBaseModel with a new object
     *
     * @param newOnenoteEntityBaseModel the object to create/update
     * @return the created OnenoteEntityBaseModel
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public OnenoteEntityBaseModel put(@Nonnull final OnenoteEntityBaseModel newOnenoteEntityBaseModel) throws ClientException {
        return send(HttpMethod.PUT, newOnenoteEntityBaseModel);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
     @Nonnull
     public OnenoteEntityBaseModelRequest select(@Nonnull final String value) {
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
     public OnenoteEntityBaseModelRequest expand(@Nonnull final String value) {
         addExpandOption(value);
         return this;
     }

}

