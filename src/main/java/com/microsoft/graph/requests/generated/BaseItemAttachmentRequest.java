// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.generated;

import com.microsoft.graph.concurrency.*;
import com.microsoft.graph.core.*;
import com.microsoft.graph.models.extensions.*;
import com.microsoft.graph.models.generated.*;
import com.microsoft.graph.http.*;
import com.microsoft.graph.requests.extensions.*;
import com.microsoft.graph.requests.generated.*;
import com.microsoft.graph.options.*;
import com.microsoft.graph.serializer.*;

import java.util.Arrays;
import java.util.EnumSet;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Base Item Attachment Request.
 */
public class BaseItemAttachmentRequest extends BaseRequest implements IBaseItemAttachmentRequest {

    /**
     * The request for the ItemAttachment
     *
     * @param requestUrl The request url
     * @param client The service client
     * @param requestOptions The options for this request
     * @param responseClass The class of the reponse
     */
    public BaseItemAttachmentRequest(final String requestUrl,
            final IBaseClient client,
            final java.util.List<? extends Option> requestOptions,
            final Class responseClass) {
        super(requestUrl, client, requestOptions, responseClass);
    }

    /**
     * Gets the ItemAttachment from the service
     * @param callback The callback to be called after success or failure.
     */
    public void get(final ICallback<ItemAttachment> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the ItemAttachment from the service
     * @return The ItemAttachment from the request.
     * @throws ClientException This exception occurs if the request was unable to complete for any reason.
     */
    public ItemAttachment get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service.
     * @param callback The callback when the deletion action has completed
     */
    public void delete(final ICallback<Void> callback) {{
        send(HttpMethod.DELETE, callback, null);
    }}

    /**
     * Delete this item from the service.
     * @throws ClientException if there was an exception during the delete operation
     */
    public void delete() throws ClientException {{
        send(HttpMethod.DELETE, null);
    }}

    /**
     * Patches this ItemAttachment with a source
     * @param sourceItemAttachment The source object with updates
     * @param callback The callback to be called after success or failure.
     */
    public void patch(final ItemAttachment sourceItemAttachment, final ICallback<ItemAttachment> callback) {
        send(HttpMethod.PATCH, callback, sourceItemAttachment);
    }

    /**
     * Patches this ItemAttachment with a source
     * @param sourceItemAttachment The source object with updates
     * @return The updated ItemAttachment
     * @throws ClientException This exception occurs if the request was unable to complete for any reason.
     */
    public ItemAttachment patch(final ItemAttachment sourceItemAttachment) throws ClientException {
        return send(HttpMethod.PATCH, sourceItemAttachment);
    }

    /**
     * Creates a ItemAttachment with a new object
     * @param newItemAttachment The new object to create
     * @param callback The callback to be called after success or failure.
     */
    public void post(final ItemAttachment newItemAttachment, final ICallback<ItemAttachment> callback) {
        send(HttpMethod.POST, callback, newItemAttachment);
    }

    /**
     * Creates a ItemAttachment with a new object
     * @param newItemAttachment The new object to create
     * @return The created ItemAttachment
     * @throws ClientException This exception occurs if the request was unable to complete for any reason.
     */
    public ItemAttachment post(final ItemAttachment newItemAttachment) throws ClientException {
        return send(HttpMethod.POST, newItemAttachment);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value The select clause
     * @return The updated request
     */
     public IItemAttachmentRequest select(final String value) {
         getQueryOptions().add(new QueryOption("$select", value));
         return (ItemAttachmentRequest)this;
     }

    /**
     * Sets the expand clause for the request
     *
     * @param value The expand clause
     * @return The updated request
     */
     public IItemAttachmentRequest expand(final String value) {
         getQueryOptions().add(new QueryOption("$expand", value));
         return (ItemAttachmentRequest)this;
     }

}

