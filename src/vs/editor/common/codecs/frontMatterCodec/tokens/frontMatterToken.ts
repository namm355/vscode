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
	 * Reference to the list of tokens that this token is made of.
	 */
	public abstract readonly tokens: readonly TSimpleDecoderToken[];

	public override get text(): string {
		return BaseToken.render(this.tokens);
	}

	// TODO: @legomushroom - extend a "composite" token
	// TODO: @legomushroom - unit test?
	public override equals(other: BaseToken): other is typeof this {
		if (super.equals(other) === false) {
			return false;
		}

		if (this.tokens.length !== other.tokens.length) {
			return false;
		}

		for (let i = 0; i < this.tokens.length; i++) {
			if (this.tokens[i].equals(other.tokens[i]) === false) {
				return false;
			}
		}

		return true;
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
