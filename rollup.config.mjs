import typescript from '@rollup/plugin-typescript';
import nodeExternals from 'rollup-plugin-node-externals';
import dts from 'rollup-plugin-dts';

const input = 'src/docx-preview.ts';

const umdGlobals = {
	jszip: 'JSZip',
	konva: 'Konva',
	'lodash-es': '_',
};

export default [
	{
		input,
		output: [
			{
				file: 'dist/docx-vellum.mjs',
				format: 'es',
				sourcemap: true,
			},
			{
				file: 'dist/docx-vellum.cjs',
				format: 'cjs',
				exports: 'named',
				sourcemap: true,
			},
			{
				// Browser/global build used by the Playwright test harness.
				file: 'dist/docx-vellum.umd.js',
				format: 'umd',
				name: 'docx',
				globals: umdGlobals,
				sourcemap: true,
			},
			{
				// Same UMD build, served by the demo pages under docs/.
				file: 'docs/js/docx-vellum.js',
				format: 'umd',
				name: 'docx',
				globals: umdGlobals,
				sourcemap: true,
			},
		],
		plugins: [
			nodeExternals(),
			typescript(),
		],
	},
	{
		input,
		output: {
			file: 'dist/docx-vellum.d.ts',
			format: 'es',
		},
		plugins: [
			dts(),
		],
	},
];
