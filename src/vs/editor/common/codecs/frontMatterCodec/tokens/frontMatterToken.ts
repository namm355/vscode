/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { BaseToken } from '../../baseToken.js';
import { TSimpleDecoderToken } from '../../simpleCodec/simpleDecoder.js';

/**
 * Base class for all tokens inside a Front Matter header.
 */
export abstract class FrontMatterToken extends BaseToken {
	/**
	 * TODO: @legomushroom
	 */
	public abstract readonly tokens: readonly TSimpleDecoderToken[];

	/**
	 * TODO: @legomushroom
	 */
	public override get text(): string {
		return BaseToken.render(this.tokens);
	}
}

/**
 * List of all currently supported value types.
 */
export type TValueTypeName = 'quoted-string' | 'boolean' | 'array' | string;

/**
 * Base class for all tokens that represent a `value` inside a Front Matter header.
 */
export abstract class FrontMatterValueToken<
	TTypeName extends TValueTypeName = TValueTypeName,
> extends FrontMatterToken {
	/**
	 * Type name of the `value` represented by this token.
	 */
	public abstract readonly valueTypeName: TTypeName;
}
