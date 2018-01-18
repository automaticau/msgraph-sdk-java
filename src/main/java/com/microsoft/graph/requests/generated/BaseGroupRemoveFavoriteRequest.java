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
 * The class for the Base Group Remove Favorite Request.
 */
public class BaseGroupRemoveFavoriteRequest extends BaseRequest implements IBaseGroupRemoveFavoriteRequest {

    /**
     * The request for this GroupRemoveFavorite
     *
     * @param requestUrl The request url
     * @param client The service client
     * @param requestOptions The options for this request
     */
    public BaseGroupRemoveFavoriteRequest(final String requestUrl, final IBaseClient client, final java.util.List<? extends Option> requestOptions) {
        super(requestUrl, client, requestOptions, Void.class);
    }

    /**
     * Creates the GroupRemoveFavorite
     *
     * @param callback The callback to be called after success or failure.
     */
    public void post(final ICallback<Void> callback) {
        final IExecutors executors = getClient().getExecutors();
        executors.performOnBackground(new Runnable() {
           @Override
           public void run() {
                try {
                    post();
                    executors.performOnForeground((Void)null, callback);
                } catch (final ClientException e) {
                    executors.performOnForeground(e, callback);
                }
           }
        });
    }

    /**
     * Creates the GroupRemoveFavorite
     *
     * @throws ClientException An exception occurs if there was an error while the request was sent.
     */
    public void post() throws ClientException {
        this.send(HttpMethod.POST, null);
    }
}
