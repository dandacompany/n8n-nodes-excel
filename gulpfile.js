const gulp = require('gulp');

function copyIcons() {
    return gulp.src('nodes/**/*.svg')
        .pipe(gulp.dest('dist/nodes'));
}

gulp.task('build:icons', copyIcons); 