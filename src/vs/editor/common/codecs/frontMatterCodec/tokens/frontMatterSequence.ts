/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { BaseToken } from '../../baseToken.js';
import { FrontMatterValueToken } from './frontMatterToken.js';
import { Word, SpacingToken } from '../../simpleCodec/tokens/index.js';
import { type TSimpleDecoderToken } from '../../simpleCodec/simpleDecoder.js';

/**
 * Token represents a generic sequence of tokens in a Front Matter header.
 */
export class FrontMatterSequence extends FrontMatterValueToken<string> {
	/**
	 * @override Because this token represent a generic sequence of tokens,
	 * the type name is represented by the text of sequence itself.
	 */
	public override get valueTypeName(): string {
		return this.text;
	}

	/**
	 * TODO: @legomushroom
	 */
	public override get tokens(): readonly TSimpleDecoderToken[] {
		return this.currentTokens;
	}

	/**
	 * TODO: @legomushroom
	 */
	private readonly currentTokens: TSimpleDecoderToken[];

	constructor(
		tokens: readonly TSimpleDecoderToken[],
	) {
		super(BaseToken.fullRange(tokens));

		this.currentTokens = [...tokens];
	}

	/**
	 * TODO: @legomushroom
	 */
	// TODO: @legomushroom - trim spaces?
	public get cleanText(): string {
		return BaseToken.render(this.tokens);
	}

	/**
	 * TODO: @legomushroom
	 */
	// TODO: @legomushroom - unit test
	public trimEnd(): readonly SpacingToken[] {
		const trimmedTokens = [];

		let index = this.currentTokens.length - 1;
		while (index >= 0) {
			const token = this.currentTokens[index];

			if (token instanceof SpacingToken) {
				trimmedTokens.push(token);
				index--;

				continue;
			}

			break;
		}

		// TODO: @legomushroom
		this.currentTokens.length = index + 1;
		if (this.currentTokens.length === 0) {
			this.collapseRangeToStart();
			this.currentTokens.push(
				new Word(this.range, ''),
			);
		}

		// TODO: @legomushroom
		this.withRange(
			BaseToken.fullRange(this.currentTokens),
		);

		return trimmedTokens;
	}

	public override toString(): string {
		return `front-matter-sequence(${this.shortText()})${this.range}`;
	}
}
