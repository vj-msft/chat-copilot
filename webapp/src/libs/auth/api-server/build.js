/* eslint-disable @typescript-eslint/no-unsafe-call */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */
import { build } from 'esbuild';
build({
        entryPoints: ['server.js'],
        bundle: true,
        platform: 'node',
        outfile: 'dist/index.js',
    })
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    .then((r) => {
        console.log(`Build succeeded.`);
    })
    .catch((e) => {
        console.log('Error building:', e.message);
        process.exit(1);
    });
