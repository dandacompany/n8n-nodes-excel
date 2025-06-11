const gulp = require('gulp');

function copyAssets() {
    return gulp.src('nodes/**/*.{png,svg}')
        .pipe(gulp.dest('dist'));
}

gulp.task('build:icons', copyAssets); 