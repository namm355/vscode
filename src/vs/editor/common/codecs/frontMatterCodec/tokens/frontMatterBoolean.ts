/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { Range } from '../../../core/range.js';
import { BaseToken } from '../../baseToken.js';
import { Word } from '../../simpleCodec/tokens/index.js';
import { FrontMatterValueToken } from './frontMatterToken.js';
import { assertDefined } from '../../../../../base/common/types.js';

/**
 * Token that represents a `boolean` value in a Front Matter header.
 */
export class FrontMatterBoolean extends FrontMatterValueToken<'boolean'> {
	/**
	 * Name of the `boolean` value type.
	 */
	public override readonly valueTypeName = 'boolean';

	constructor(
		range: Range,
		public readonly value: boolean,
	) {
		super(range);
	}

	/**
	 * TODO: @legomushroom
	 */
	public static fromToken(token: Word): FrontMatterBoolean {
		const value = asBoolean(token);

		assertDefined(
			value,
			`Cannot convert '${token}' to a boolean value.`,
		);

		return new FrontMatterBoolean(token.range, value);
	}

	/**
	 * TODO: @legomushroom
	 */
	public static tryFromToken(
		token: BaseToken,
	): FrontMatterBoolean | null {
		if (token instanceof Word === false) {
			return null;
		}

		try {
			return FrontMatterBoolean.fromToken(token);
		} catch (_error) {
			// noop
			return null;
		}
	}

	public override get text(): string {
		return `${this.value}`;
	}

	public override toString(): string {
		return `front-matter-boolean(${this.shortText()})${this.range}`;
	}
}

/**
 * Try to convert a {@link Word} token to a `boolean` value.
 */
const asBoolean = (
	token: Word,
): boolean | null => {
	if (token.text.toLowerCase() === 'true') {
		return true;
	}

	if (token.text.toLowerCase() === 'false') {
		return false;
	}

	return null;
};
