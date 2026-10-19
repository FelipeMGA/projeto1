from pathlib import Path
import sys

root = Path(sys.argv[1] if len(sys.argv) > 1 else 'upstream')
p = root / 'android/app/src/main/java/com/gevr/port/TouchOverlayView.java'
s = p.read_text()


def rep(old, new, n=1):
    global s
    c = s.count(old)
    if c < n:
        raise SystemExit(f'missing patch target ({c}<{n}): {old[:120]!r}')
    s = s.replace(old, new, n)


rep(' * v7 input model:', ' * v8 input model:')
rep(
    ' * - touch aim has its own sensitivity;\n * - the phone gyroscope can drive the camera always or only while aiming;\n * - touch, gyro and a physical SDL gamepad are merged by the native layer.',
    ' * - touch aim has its own sensitivity plus independent X/Y inversion;\n * - the phone gyroscope can drive the camera always or only while aiming, with independent X/Y inversion;\n * - gyro uses higher-rate sampling, adaptive filtering and a soft deadzone for low-latency precision;\n * - touch, gyro and a physical SDL gamepad are merged by the native layer.'
)

rep(
    '    private float gyroAimSensitivity;\n    private float gyroFilteredX, gyroFilteredY;\n    private float opacity;',
    '    private float gyroAimSensitivity;\n    private float gyroFilteredX, gyroFilteredY;\n    private boolean gyroInvertX, gyroInvertY;\n    private boolean touchInvertX, touchInvertY;\n    private float opacity;'
)

rep(
    '    private final RectF gyroAimMinus = new RectF(), gyroAimPlus = new RectF();\n    private final RectF scaleMinus = new RectF(), scalePlus = new RectF();',
    '    private final RectF gyroAimMinus = new RectF(), gyroAimPlus = new RectF();\n    private final RectF gyroInvXToggle = new RectF(), gyroInvYToggle = new RectF();\n    private final RectF touchInvXToggle = new RectF(), touchInvYToggle = new RectF();\n    private final RectF scaleMinus = new RectF(), scalePlus = new RectF();'
)

rep(
    '        gyroAimSensitivity = prefs.getFloat("gyroAimSensitivity", 0.85f);\n        opacity = prefs.getFloat("opacity", 0.54f);',
    '        gyroAimSensitivity = prefs.getFloat("gyroAimSensitivity", 0.85f);\n        // v8 default fixes the horizontal direction reported on the v7 test build.\n        gyroInvertX = prefs.getBoolean("gyroInvertX", false);\n        gyroInvertY = prefs.getBoolean("gyroInvertY", false);\n        touchInvertX = prefs.getBoolean("touchInvertX", false);\n        touchInvertY = prefs.getBoolean("touchInvertY", false);\n        opacity = prefs.getFloat("opacity", 0.54f);'
)

rep(
    '            gyroRegistered = sensorManager.registerListener(this, gyroSensor, SensorManager.SENSOR_DELAY_GAME);',
    '            // Request ~100 Hz to reduce gyro latency.\n            gyroRegistered = sensorManager.registerListener(this, gyroSensor, 10000);'
)

old_layout = '''            float left = l + pw*.04f, right = l + pw*.54f;
            float col = pw*.42f, pairGap = pw*.025f;
            float rowH = ph*.115f;
            float y1=t+ph*.165f, y2=t+ph*.330f, y3=t+ph*.495f, y4=t+ph*.660f;
            makePair(sensMinus,sensPlus,left,y1,col,rowH,pairGap);
            makePair(aimSensMinus,aimSensPlus,right,y1,col,rowH,pairGap);
            gyroModeToggle.set(left,y2,left+col,y2+rowH);
            makePair(gyroSensMinus,gyroSensPlus,right,y2,col,rowH,pairGap);
            makePair(gyroAimMinus,gyroAimPlus,left,y3,col,rowH,pairGap);
            makePair(scaleMinus,scalePlus,right,y3,col,rowH,pairGap);
            makePair(alphaMinus,alphaPlus,left,y4,col,rowH,pairGap);
            padToggle.set(right,y4,right+col,y4+rowH);'''
new_layout = '''            float left = l + pw*.04f, right = l + pw*.54f;
            float col = pw*.42f, pairGap = pw*.025f;
            float rowH = ph*.075f;
            float y1=t+ph*.150f, y2=t+ph*.285f, y3=t+ph*.420f;
            float y4=t+ph*.555f, y5=t+ph*.690f, y6=t+ph*.825f;
            makePair(sensMinus,sensPlus,left,y1,col,rowH,pairGap);
            makePair(aimSensMinus,aimSensPlus,right,y1,col,rowH,pairGap);
            gyroModeToggle.set(left,y2,left+col,y2+rowH);
            makePair(gyroSensMinus,gyroSensPlus,right,y2,col,rowH,pairGap);
            makePair(gyroAimMinus,gyroAimPlus,left,y3,col,rowH,pairGap);
            gyroInvXToggle.set(right,y3,right+col,y3+rowH);
            gyroInvYToggle.set(left,y4,left+col,y4+rowH);
            touchInvXToggle.set(right,y4,right+col,y4+rowH);
            touchInvYToggle.set(left,y5,left+col,y5+rowH);
            makePair(scaleMinus,scalePlus,right,y5,col,rowH,pairGap);
            makePair(alphaMinus,alphaPlus,left,y6,col,rowH,pairGap);
            padToggle.set(right,y6,right+col,y6+rowH);'''
rep(old_layout, new_layout)

rep('        float dead = .09f;', '        float dead = .065f;')
rep(
    '        else if (mag > 0) { float scaled=(mag-dead)/(1-dead)/mag; nx*=scaled; ny*=scaled; }',
    '        else if (mag > 0) {\n            float shaped=(mag-dead)/(1-dead);\n            shaped=(float)Math.pow(clamp(shaped,0f,1f), 1.04f);\n            float scaled=shaped/mag; nx*=scaled; ny*=scaled;\n        }'
)

old_look = '''        float sens = aiming() ? aimSensitivity : touchSensitivity;
        // Relative motion, not distance from a virtual stick centre. Smaller base = faster turn.
        float base = Math.max(dp(26), Math.min(getWidth(), getHeight()) * .040f);
        lookImpulseX = clamp(lookImpulseX + (dx/base) * sens, -3.0f, 3.0f);
        lookImpulseY = clamp(lookImpulseY + (-dy/base) * sens, -3.0f, 3.0f);
        ensureLookPump();'''
new_look = '''        float sens = aiming() ? aimSensitivity : touchSensitivity;
        if (touchInvertX) dx = -dx;
        if (touchInvertY) dy = -dy;
        // PUBG-like relative look: small swipes remain precise while fast swipes get a gentle boost.
        float base = Math.max(dp(24), Math.min(getWidth(), getHeight()) * .037f);
        float travel = (float)Math.hypot(dx,dy) / base;
        float boost = 1.0f + Math.min(.70f, travel * .22f);
        if (aiming()) boost = 1.0f + Math.min(.30f, travel * .12f);
        lookImpulseX = clamp(lookImpulseX + (dx/base) * sens * boost, -4.5f, 4.5f);
        lookImpulseY = clamp(lookImpulseY + (-dy/base) * sens * boost, -4.5f, 4.5f);
        ensureLookPump();'''
rep(old_look, new_look)

rep(
    '            lookImpulseX -= touchLookX * .46f;\n            lookImpulseY -= touchLookY * .46f;',
    '            float consume = aiming() ? .50f : .39f;\n            lookImpulseX -= touchLookX * consume;\n            lookImpulseY -= touchLookY * consume;'
)

old_gyro = '''        // screenY = yaw axis, screenX = pitch axis. Deadzone filters stationary sensor noise.
        if (Math.abs(screenY) < .012f) screenY = 0f;
        if (Math.abs(screenX) < .012f) screenX = 0f;
        float sens = aiming() ? gyroAimSensitivity : gyroSensitivity;
        float targetX = clamp(-screenY * .72f * sens, -1f, 1f);
        float targetY = clamp(-screenX * .72f * sens, -1f, 1f);
        // Mild low-pass filter: PUBG-like stability without making the gyro laggy.
        gyroFilteredX += (targetX - gyroFilteredX) * .42f;
        gyroFilteredY += (targetY - gyroFilteredY) * .42f;
        gyroLookX = Math.abs(gyroFilteredX) < .004f ? 0f : gyroFilteredX;
        gyroLookY = Math.abs(gyroFilteredY) < .004f ? 0f : gyroFilteredY;'''
new_gyro = '''        // screenY = yaw, screenX = pitch. v8 corrects yaw direction by default.
        // Soft deadzone suppresses stationary noise without a visible jump at the threshold.
        float yaw = screenY;
        float pitch = -screenX;
        float dz = .0085f;
        if (Math.abs(yaw) <= dz) yaw = 0f;
        else yaw = Math.copySign(Math.abs(yaw)-dz, yaw);
        if (Math.abs(pitch) <= dz) pitch = 0f;
        else pitch = Math.copySign(Math.abs(pitch)-dz, pitch);
        if (gyroInvertX) yaw = -yaw;
        if (gyroInvertY) pitch = -pitch;
        float sens = aiming() ? gyroAimSensitivity : gyroSensitivity;
        float targetX = clamp(yaw * .80f * sens, -1f, 1f);
        float targetY = clamp(pitch * .80f * sens, -1f, 1f);
        // Adaptive low-pass: micro-corrections are stable and fast turns get lower latency.
        float motion = Math.max(Math.abs(targetX), Math.abs(targetY));
        float alpha = clamp(.24f + motion * .52f, .24f, .76f);
        if (aiming()) alpha = clamp(alpha - .04f, .20f, .72f);
        gyroFilteredX += (targetX - gyroFilteredX) * alpha;
        gyroFilteredY += (targetY - gyroFilteredY) * alpha;
        gyroLookX = Math.abs(gyroFilteredX) < .0025f ? 0f : gyroFilteredX;
        gyroLookY = Math.abs(gyroFilteredY) < .0025f ? 0f : gyroFilteredY;'''
rep(old_gyro, new_gyro)

old_handler = '''                        else if (hit(gyroAimMinus,x,y)) gyroAimSensitivity=clamp(gyroAimSensitivity-.10f,.10f,3f);
                        else if (hit(gyroAimPlus,x,y)) gyroAimSensitivity=clamp(gyroAimSensitivity+.10f,.10f,3f);
                        else if (hit(alphaMinus,x,y)) opacity=clamp(opacity-.08f,.15f,.90f);'''
new_handler = '''                        else if (hit(gyroAimMinus,x,y)) gyroAimSensitivity=clamp(gyroAimSensitivity-.10f,.10f,4f);
                        else if (hit(gyroAimPlus,x,y)) gyroAimSensitivity=clamp(gyroAimSensitivity+.10f,.10f,4f);
                        else if (hit(gyroInvXToggle,x,y)) gyroInvertX=!gyroInvertX;
                        else if (hit(gyroInvYToggle,x,y)) gyroInvertY=!gyroInvertY;
                        else if (hit(touchInvXToggle,x,y)) touchInvertX=!touchInvertX;
                        else if (hit(touchInvYToggle,x,y)) touchInvertY=!touchInvertY;
                        else if (hit(alphaMinus,x,y)) opacity=clamp(opacity-.08f,.15f,.90f);'''
rep(old_handler, new_handler)
rep(
    '                        else if (hit(gyroSensMinus,x,y)) gyroSensitivity=clamp(gyroSensitivity-.10f,.10f,3f);\n                        else if (hit(gyroSensPlus,x,y)) gyroSensitivity=clamp(gyroSensitivity+.10f,.10f,3f);',
    '                        else if (hit(gyroSensMinus,x,y)) gyroSensitivity=clamp(gyroSensitivity-.10f,.10f,4f);\n                        else if (hit(gyroSensPlus,x,y)) gyroSensitivity=clamp(gyroSensitivity+.10f,.10f,4f);'
)

rep(
    '                .putFloat("gyroAimSensitivity",gyroAimSensitivity)\n                .putFloat("opacity",opacity).putFloat("scale",controlScale)',
    '                .putFloat("gyroAimSensitivity",gyroAimSensitivity)\n                .putBoolean("gyroInvertX",gyroInvertX).putBoolean("gyroInvertY",gyroInvertY)\n                .putBoolean("touchInvertX",touchInvertX).putBoolean("touchInvertY",touchInvertY)\n                .putFloat("opacity",opacity).putFloat("scale",controlScale)'
)

rep(
    '    private String gyroModeName() {\n        if (gyroSensor == null) return "SEM SENSOR";\n        return gyroMode==GYRO_OFF ? "DESLIGADO" : (gyroMode==GYRO_AIM ? "SÓ MIRA" : "SEMPRE");\n    }',
    '    private String gyroModeName() {\n        if (gyroSensor == null) return "SEM SENSOR";\n        return gyroMode==GYRO_OFF ? "DESLIGADO" : (gyroMode==GYRO_AIM ? "SÓ MIRA" : "SEMPRE");\n    }\n    private String axisName(boolean inverted) { return inverted ? "INVERTIDO" : "NORMAL"; }'
)

old_draw = '''            label(c,"CÂMERA TOUCH  " + String.format("%.2f",touchSensitivity),l+pw*.25f,t+ph*.145f);
            label(c,"MIRA TOUCH  " + String.format("%.2f",aimSensitivity),l+pw*.75f,t+ph*.145f);
            mini(c,sensMinus,"−"); mini(c,sensPlus,"+"); mini(c,aimSensMinus,"−"); mini(c,aimSensPlus,"+");

            label(c,"GIROSCÓPIO",l+pw*.25f,t+ph*.310f);
            label(c,"SENS. GIRO  " + String.format("%.2f",gyroSensitivity),l+pw*.75f,t+ph*.310f);
            mini(c,gyroModeToggle,gyroModeName()); mini(c,gyroSensMinus,"−"); mini(c,gyroSensPlus,"+");

            label(c,"GIRO MIRA  " + String.format("%.2f",gyroAimSensitivity),l+pw*.25f,t+ph*.475f);
            label(c,"TAMANHO",l+pw*.75f,t+ph*.475f);
            mini(c,gyroAimMinus,"−"); mini(c,gyroAimPlus,"+"); mini(c,scaleMinus,"−"); mini(c,scalePlus,"+");

            label(c,"OPACIDADE",l+pw*.25f,t+ph*.640f);
            label(c,"GAMEPAD",l+pw*.75f,t+ph*.640f);
            mini(c,alphaMinus,"−"); mini(c,alphaPlus,"+"); mini(c,padToggle,hideOnGamepad?"AUTO-HIDE":"SEMPRE");'''
new_draw = '''            label(c,"CÂMERA TOUCH  " + String.format("%.2f",touchSensitivity),l+pw*.25f,t+ph*.135f);
            label(c,"MIRA TOUCH  " + String.format("%.2f",aimSensitivity),l+pw*.75f,t+ph*.135f);
            mini(c,sensMinus,"−"); mini(c,sensPlus,"+"); mini(c,aimSensMinus,"−"); mini(c,aimSensPlus,"+");

            label(c,"GIROSCÓPIO",l+pw*.25f,t+ph*.270f);
            label(c,"SENS. GIRO  " + String.format("%.2f",gyroSensitivity),l+pw*.75f,t+ph*.270f);
            mini(c,gyroModeToggle,gyroModeName()); mini(c,gyroSensMinus,"−"); mini(c,gyroSensPlus,"+");

            label(c,"GIRO MIRA  " + String.format("%.2f",gyroAimSensitivity),l+pw*.25f,t+ph*.405f);
            label(c,"GIRO X",l+pw*.75f,t+ph*.405f);
            mini(c,gyroAimMinus,"−"); mini(c,gyroAimPlus,"+"); mini(c,gyroInvXToggle,axisName(gyroInvertX));

            label(c,"GIRO Y",l+pw*.25f,t+ph*.540f);
            label(c,"TOUCH X",l+pw*.75f,t+ph*.540f);
            mini(c,gyroInvYToggle,axisName(gyroInvertY)); mini(c,touchInvXToggle,axisName(touchInvertX));

            label(c,"TOUCH Y",l+pw*.25f,t+ph*.675f);
            label(c,"TAMANHO",l+pw*.75f,t+ph*.675f);
            mini(c,touchInvYToggle,axisName(touchInvertY)); mini(c,scaleMinus,"−"); mini(c,scalePlus,"+");

            label(c,"OPACIDADE",l+pw*.25f,t+ph*.810f);
            label(c,"GAMEPAD",l+pw*.75f,t+ph*.810f);
            mini(c,alphaMinus,"−"); mini(c,alphaPlus,"+"); mini(c,padToggle,hideOnGamepad?"AUTO-HIDE":"SEMPRE");'''
rep(old_draw, new_draw)

p.write_text(s)
print('v8 patch applied: corrected gyro X + axis toggles + adaptive gyro + improved touch look')
